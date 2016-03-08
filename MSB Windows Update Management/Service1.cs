using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using WUApiLib;
using System.Management;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;

namespace MSB_Windows_Update_Management
{
    public partial class MSBWindowsUpdateManagement : ServiceBase
    {
        private System.ComponentModel.IContainer components = null;
        private System.Diagnostics.EventLog eventLog1;

        public dynamic appSettings;

        public ISystemInformation uSysInfo = new SystemInformation();
        public UpdateSession uSession = new UpdateSession();
        public IUpdateSearcher uSearcher;
        public ISearchResult uResult;
        public UpdateDownloader uDownloader;
        public IUpdateInstaller uInstaller;
        public UpdateCollection uDownloadList;
        public UpdateCollection uInstallList;
        public AutomaticUpdatesClass uAutomaticUpdate = new AutomaticUpdatesClass();
        public bool ShouldInstallUpdates;

        public string connectionString = "Server=msbsqlrpt;Database=ITDashboard;User Id=ITDashboardUser;Password=ITDashboardUser;";

        public string countUpdatesForServer = "SELECT count(*) as count " + 
            "FROM [dbo].[update_details] as ud " + 
            "INNER JOIN dbo.updates as u " + 
            "ON u.id = ud.update_id " + 
            "INNER JOIN dbo.servers as s " + 
            "ON s.id = ud.server_id " + 
            "WHERE u.title = @title " + 
            "AND s.name = @hostname";

        public string countUpdates = "SELECT count(*) as count FROM [dbo].[updates] WHERE title = @title";

        public string insertUpdate = "INSERT INTO [dbo].[updates] (title, description, kb_article, created_at, updated_at) VALUES (@title, @description, @kb_article, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

        public string insertUpdateForServer = "INSERT INTO [dbo].[update_details] (update_id, server_id, eula_accepted, downloaded_flag, hidden_flag, installed_flag, mandatory_flag, max_download_size, min_download_size, approved_flag, created_at, updated_at) VALUES ( " +
            "@update_id, " +
            "@server_id, " +
            "@eula_accepted, " +
            "@downloaded_flag, " +
            "@hidden_flag, " +
            "@installed_flag, " +
            "@mandatory_flag, " +
            "@max_download_size, " +
            "@min_download_size, " +
            "0, " +
            "CAST(SYSDATETIME() AS VARCHAR(19)), " +
            "CAST(SYSDATETIME() AS VARCHAR(19)) )";

        public string insertServer = "INSERT INTO dbo.servers (name, ip, operating_system_id, created_at, updated_at) VALUES (@hostname, @ip, 1, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

        public string updateServer = "UPDATE dbo.servers SET " +
            "last_windows_update = @last_windows_update, " +
            "updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) " + 
            "WHERE name = @hostname";

        public string updateUpdateForServer = "UPDATE [dbo].[update_details] SET " +
            "eula_accepted = @eula_accepted, " +
            "downloaded_flag = @downloaded_flag, " +
            "hidden_flag = @hidden_flag, " +
            "installed_flag = @installed_flag, " +
            "mandatory_flag = @mandatory_flag, " +
            "updated_at = CAST(SYSDATETIME() AS VARCHAR(19)) " +
            "WHERE server_id = @server_id AND update_id = @update_id";

        public string queryGetUpdateId = "SELECT id FROM dbo.updates WHERE title = @title";

        public string queryGetServerId = "SELECT id FROM dbo.servers WHERE name = @hostname";

        public string queryGetServerStatus = "SELECT status FROM dbo.servers WHERE name = @hostname";

        public string setServerStatus = "UPDATE dbo.servers SET status=@status, updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) WHERE name = @hostname";

        public string queryUpdateApproved = "SELECT count(id) as count FROM dbo.update_details WHERE installed_flag=0 AND approved_flag=1 AND server_id=@server_id AND update_id=@update_id";

        public string queryUpdateHidden = "SELECT count(id) as count FROM dbo.update_details WHERE hidden_flag=1 AND server_id=@server_id AND update_id=@update_id";

        // Set up a timer to look for updates
        public System.Timers.Timer lookForUpdatesTimer = new System.Timers.Timer();

        public System.Timers.Timer updateServerInfoTimer = new System.Timers.Timer();

        public System.Timers.Timer getServerStatusTimer = new System.Timers.Timer();

        public MSBWindowsUpdateManagement()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("Updates"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "Updates", "MSB");
            }
            eventLog1.Source = "Updates";
            eventLog1.Log = "MSB";

            if (Environment.UserInteractive)
            {
                this.LookForUpdates();
            }
        }

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry("Started Service MSB Windows Update Management");


            this.SetupTimers();
           

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            //this.LookForUpdates();
        }

        private void SetupTimers()
        {
            // Check for available updates timer    
            lookForUpdatesTimer.Interval = 12000; //1000*60*30; // 30 min
            lookForUpdatesTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnLookForUpdatesTimer);
            lookForUpdatesTimer.Start();

            // Update Server Info timer    
            updateServerInfoTimer.Interval = 5000; //1000*60*30; // 30 min
            updateServerInfoTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnUpdateServerInfoTimer);
            updateServerInfoTimer.Start();

            // Update Server Info timer    
            getServerStatusTimer.Interval = 1000; //1000*60*30; // 30 min
            getServerStatusTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnGetServerStatusTimer);
            getServerStatusTimer.Start();
        }

        public void OnLookForUpdatesTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            lookForUpdatesTimer.Interval = 1000 * 60 * 60 * 8; // 8 hours
            this.LookForUpdates();    
        }

        public void OnUpdateServerInfoTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            updateServerInfoTimer.Interval = 1000 * 60 * 60 * 8; // 8 hours
            this.UpdateServerInfo();
        }

        public void OnGetServerStatusTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            getServerStatusTimer.Interval = 1000 * 60 * 2; // 2 mins
            
            switch ( this.GetServerStatus() )
            {
                case "Ready For Updates" :
                case "Look For Updates" :
                    this.LookForUpdates();
                    break;

                case "Ready For Reboot" :
                    this.RebootComputer();
                    break;
            }    
        }


        protected override void OnStop()
        {
        }

        public void UpdateServerInfo()
        {
            DateTime? lastUpdateSuccess = null;
            if ( uAutomaticUpdate.Results.LastInstallationSuccessDate is DateTime)
            {
                lastUpdateSuccess = new DateTime( ( (DateTime)uAutomaticUpdate.Results.LastInstallationSuccessDate).Ticks, DateTimeKind.Local);
            }

            if ( lastUpdateSuccess != null )
            {
                this.SetServerLastWindowsUpdate(lastUpdateSuccess);
            }

        }

        public string GetIPAddress()
        {

            StringBuilder sb = new StringBuilder(); 

            // Get a list of all network interfaces (usually one per network card, dialup, and VPN connection) 
            NetworkInterface[] networkInterfaces = NetworkInterface.GetAllNetworkInterfaces(); 

            foreach (NetworkInterface network in networkInterfaces) 
            { 
                // Read the IP configuration for each network 
                IPInterfaceProperties properties = network.GetIPProperties(); 

                // Each network interface may have multiple IP addresses 
                foreach (IPAddressInformation address in properties.UnicastAddresses) 
                { 
                    // We're only interested in IPv4 addresses for now 
                    if (address.Address.AddressFamily != AddressFamily.InterNetwork) 
                        continue; 

                    // Ignore loopback addresses (e.g., 127.0.0.1) 
                    if (IPAddress.IsLoopback(address.Address)) 
                        continue; 

                    sb.AppendLine(address.Address.ToString() + " (" + network.Name + ")"); 
                } 
            } 

            return sb.ToString(); 
        }

        public void storeServerInDashboard()
        {
            SqlCommand command = new SqlCommand(insertServer);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            command.Parameters.AddWithValue("@ip", this.GetIPAddress());
            this.executeNonQuery(command);
        }

        public void SetServerLastWindowsUpdate(DateTime? lastUpdateSuccess)
        {
            SqlCommand command = new SqlCommand(updateServer);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            command.Parameters.AddWithValue("@last_windows_update", lastUpdateSuccess);
            this.executeNonQuery(command);
        }

        public void SetServerStatus(string status)
        {
            SqlCommand command = new SqlCommand(setServerStatus);
            command.Parameters.AddWithValue("@status", status);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            this.executeNonQuery(command);
        }

        public string GetServerStatus()
        {
            SqlCommand command = new SqlCommand(queryGetServerStatus);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = this.executeQuery(command);

            if ( rows.Rows.Count < 1)
            {
                return "";
            }

            DataRow row = rows.Rows[0];

            return row["status"].ToString();
        }

        private void status(string message, EventLogEntryType type = EventLogEntryType.Information)
        {
            eventLog1.WriteEntry(message, type);
            Console.WriteLine(message);
            this.SetServerStatus(message);
        }



        public void LookForUpdates()
        {
            ShouldInstallUpdates = (this.GetServerStatus() == "Ready For Updates");

            uDownloadList = new UpdateCollection();
            uInstallList = new UpdateCollection();
            uSearcher = uSession.CreateUpdateSearcher();
            List<String> updateList = new List<string>();


            this.status("Looking For Available Updates...");

            uSearcher.Online = true;
            uResult = uSearcher.Search("IsInstalled=0 and Type='Software'");

            this.status("There are " + uResult.Updates.Count + " available updates");


            updateList.Add("Available Updates...");
            foreach (IUpdate update in uResult.Updates)
            {
                this.processUpdate(update, updateList, uDownloadList, uInstallList);
            }
            
            // download any missing updates
            this.downloadUpdates(uDownloadList);
            
            // install any queued updates
            this.installUpdates(uInstallList);

            // update the status of updates
            this.updateUpdates();

            // update server info
            this.UpdateServerInfo();

            eventLog1.WriteEntry(string.Join(System.Environment.NewLine, updateList), EventLogEntryType.Information);
            Console.WriteLine(string.Join(System.Environment.NewLine, updateList));

            if (uSysInfo.RebootRequired)
            {                          
                this.status("Reboot Required");
            }
            else
            {
                this.status("Idle");
            }
            
        }

        private void processUpdate(IUpdate update, List<String> updateList, UpdateCollection uDownloadList, UpdateCollection uInstallList)
        {
            
            this.storeUpdateInDashboard(update);

            this.updateUpdateInDashboardForThisServer(update);

            if (!update.IsDownloaded)
            {
                uDownloadList.Add(update);
            }

            if (this.isUpdateApproved(update))
            {
                uInstallList.Add(update);
            }

            updateList.Add(update.Title.ToString());
        }

        private void updateUpdates()
        {
            // TODO: Insert monitoring activities here.
            this.status("Updating updates...");

            uSearcher = uSession.CreateUpdateSearcher();
            uSearcher.Online = true;
            uResult = uSearcher.Search("Type='Software'");

            foreach (IUpdate update in uResult.Updates)
            {
                update.IsHidden = this.isUpdateHidden(update);
                
                if ( this.isUpdateInDashboardForThisServer(update))
                {
                    eventLog1.WriteEntry("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title, EventLogEntryType.Information);
                    Console.WriteLine("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title);
                    this.updateUpdateInDashboardForThisServer(update);
                }
            }
        }

        private string getKBArticle(string title)
        {
            string result = "";
            string pattern = @"(KB\d+)";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            MatchCollection matches = regex.Matches(title);

            if (matches.Count > 0)
            {
                result = matches[0].Value;
            }
            return result;
        }

        private dynamic storeUpdateInDashboardForThisServer(IUpdate update)
        {            
                    
            // first get update_id
            int update_id = this.getUpdateId(update);

            if (update_id < 1)
            {
                // update not in dashboard
                this.storeUpdateInDashboard(update);
                return this.storeUpdateInDashboardForThisServer(update);
            }

            // get server_id
            int server_id = this.getServerId();

            if (server_id < 1)
            {
                // server is not in dashboard
                this.storeServerInDashboard();
                return this.storeUpdateInDashboardForThisServer(update);
            }
                   
            SqlCommand command = new SqlCommand(insertUpdateForServer);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            command.Parameters.AddWithValue("@eula_accepted", update.EulaAccepted);
            command.Parameters.AddWithValue("@downloaded_flag", update.IsDownloaded);
            command.Parameters.AddWithValue("@hidden_flag", update.IsHidden);
            command.Parameters.AddWithValue("@installed_flag", update.IsInstalled);
            command.Parameters.AddWithValue("@mandatory_flag", update.IsMandatory);
            command.Parameters.AddWithValue("@max_download_size", update.MaxDownloadSize);
            command.Parameters.AddWithValue("@min_download_size", update.MinDownloadSize);

            this.executeNonQuery(command);

            return true;
        }

        private dynamic updateUpdateInDashboardForThisServer(IUpdate update)
        {
            if ( ! this.isUpdateInDashboardForThisServer(update) )
            {
                return this.storeUpdateInDashboardForThisServer(update);
            }

            // first get update_id
            int update_id = this.getUpdateId(update);

            // get server_id
            int server_id = this.getServerId();


            SqlCommand command = new SqlCommand(updateUpdateForServer);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            command.Parameters.AddWithValue("@eula_accepted", update.EulaAccepted);
            command.Parameters.AddWithValue("@downloaded_flag", update.IsDownloaded);
            command.Parameters.AddWithValue("@hidden_flag", update.IsHidden);
            command.Parameters.AddWithValue("@installed_flag", update.IsInstalled);
            command.Parameters.AddWithValue("@mandatory_flag", update.IsMandatory);
            command.Parameters.AddWithValue("@max_download_size", update.MaxDownloadSize);
            command.Parameters.AddWithValue("@min_download_size", update.MinDownloadSize);

            this.executeNonQuery(command);

            return true;
        }

        private void storeUpdateInDashboard(IUpdate update)
        {
            if ( ! this.isUpdateInDashboard(update) )
            {
                SqlCommand command = new SqlCommand(insertUpdate);
                command.Parameters.AddWithValue("@title", update.Title);
                command.Parameters.AddWithValue("@description", update.Title);
                command.Parameters.AddWithValue("@kb_article", this.getKBArticle(update.Title));

                this.executeNonQuery(command);
            }    
        }

        private bool isUpdateApproved(IUpdate update)
        {
            int count;

            int update_id = this.getUpdateId(update);
            int server_id = this.getServerId();

            SqlCommand command = new SqlCommand(queryUpdateApproved);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            DataTable rows = this.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        private bool isUpdateHidden(IUpdate update)
        {
            int count;

            int update_id = this.getUpdateId(update);
            int server_id = this.getServerId();

            SqlCommand command = new SqlCommand(queryUpdateHidden);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            DataTable rows = this.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        private bool isUpdateInDashboard(IUpdate update)
        {
            int count;
            SqlCommand command = new SqlCommand(countUpdates);
            command.Parameters.AddWithValue("@title", update.Title);
            DataTable rows = this.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        private bool isUpdateInDashboardForThisServer(IUpdate update)
        {
            int count;
            SqlCommand command = new SqlCommand(countUpdatesForServer);
            command.Parameters.AddWithValue("@title", update.Title);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = this.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        private int getUpdateId(IUpdate update)
        {
            int id;
            SqlCommand command = new SqlCommand(queryGetUpdateId);
            command.Parameters.AddWithValue("@title", update.Title);
            DataTable rows = this.executeQuery(command);

            if ( rows.Rows.Count < 1 )
            {
                return 0;
            }
            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        private int getServerId()
        {
            int id;
            SqlCommand command = new SqlCommand(queryGetServerId);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = this.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return 0;
            }

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        

        private void executeNonQuery(SqlCommand command)
        {
            using (SqlConnection dbc = new SqlConnection(connectionString))
            {
                try
                {
                    if (dbc != null && dbc.State != ConnectionState.Open)
                    {
                        dbc.Close();
                        dbc.Open();
                    }

                    command.Connection = dbc;
                    
                    command.ExecuteNonQuery();

                }

                catch (SqlException e)
                {
                    eventLog1.WriteEntry(e.Message, EventLogEntryType.Error);
                }

                finally
                {
                    dbc.Close();
                }
            }
        }

        private DataTable executeQuery(SqlCommand command)
        {
            using (SqlConnection dbc = new SqlConnection(connectionString))
            {
                try
                {
                    if (dbc != null && dbc.State != ConnectionState.Open)
                    {
                        dbc.Close();
                        dbc.Open();
                    }

                    command.Connection = dbc;                    
                    
                    SqlDataReader reader = command.ExecuteReader();

                    try
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        return dt;
                    }

                    finally
                    {
                        reader.Close();
                    }
                }

                finally
                {
                    dbc.Close();
                }
            }
        }

        private void downloadUpdates(UpdateCollection updateList)
        {
            try
            {
                this.status("Downloading " + updateList.Count + " Updates");
                uDownloader = uSession.CreateUpdateDownloader();
                uDownloader.Updates = updateList;
                uDownloader.Download();
            }

            catch (COMException e)
            {
                eventLog1.WriteEntry(e.Message, EventLogEntryType.Error);
            }
        }

        private void installUpdates(UpdateCollection updateList)
        {
            

            if ((updateList.Count > 0) && ShouldInstallUpdates)
            {
                this.status("Installing Updates...");

                try
                {
                    this.status("Installing " + updateList.Count + " Updates");
                    uInstaller = uSession.CreateUpdateInstaller();
                    uInstaller.AllowSourcePrompts = false;
                    if (uInstaller.IsBusy)
                    {
                        this.status("Update Installer Busy, Try Again");
                        return;
                    }

                    if (uInstaller.RebootRequiredBeforeInstallation)
                    {
                        this.status("Reboot Required");
                        return;
                    }

                    uInstaller.Updates = updateList;
                    uInstaller.Install(); 

                }
                catch (Exception e)
                {
                    eventLog1.WriteEntry(e.Message, EventLogEntryType.Error);
                }
            }
        }

        private void RebootComputer()
        {
            this.status("Rebooting in 5 min");
            System.Diagnostics.Process.Start("shutdown.exe", "-r -t 300 -c \"Rebooting in 5. Run 'shutdown -a' to abort.\"");
        }

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public long dwServiceType;
            public ServiceState dwCurrentState;
            public long dwControlsAccepted;
            public long dwWin32ExitCode;
            public long dwServiceSpecificExitCode;
            public long dwCheckPoint;
            public long dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
    }
}
