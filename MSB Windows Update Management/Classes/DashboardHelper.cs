using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Data;
using System.Diagnostics;
using System.IO;
using WUApiLib;

namespace MSB_Windows_Update_Management
{
    public partial class DashboardHelper
    {
        public AutomaticUpdatesClass uAutomaticUpdate = new AutomaticUpdatesClass();


        public DashboardHelper()
        {

        }

        public void UpdateServerInfo()
        {
            DateTime? lastUpdateSuccess = null;
            if (uAutomaticUpdate.Results.LastInstallationSuccessDate is DateTime)
            {
                lastUpdateSuccess = new DateTime(((DateTime)uAutomaticUpdate.Results.LastInstallationSuccessDate).Ticks, DateTimeKind.Local);
            }

            if (lastUpdateSuccess != null)
            {
                this.SetServerLastWindowsUpdate(lastUpdateSuccess);
            }

        }

        public void CheckIn()
        {
            SqlCommand command = new SqlCommand(Program.DB.checkIn);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            Program.DB.executeNonQuery(command);
        }

        public void status(string message, EventLogEntryType type = EventLogEntryType.Information)
        {
            Program.Events.WriteEntry(message, type);
            Console.WriteLine(message);
            this.SetServerStatus(message);
        }

        public void storeServerInDashboard()
        {
            SqlCommand command = new SqlCommand(Program.DB.insertServer);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            command.Parameters.AddWithValue("@ip", this.GetIPAddress());
            Program.DB.executeNonQuery(command);
        }

        public void InsertUpdateBatch(string result, int batch_id)
        {
            SqlCommand command = new SqlCommand(Program.DB.insertUpdateBatch);
            command.Parameters.AddWithValue("@server_id", this.getServerId());
            command.Parameters.AddWithValue("@batch_id", batch_id);
            command.Parameters.AddWithValue("@result", result);
            Program.DB.executeNonQuery(command);
        }

        public void CleanupUpdates( List<int> update_ids )
        {
            string comm = string.Format( Program.DB.deleteSupersededUpdates, string.Join(",",update_ids) );
            SqlCommand command = new SqlCommand(comm);
            command.Parameters.AddWithValue("@server_id", this.getServerId());
            Program.DB.executeNonQuery(command);
        }

        public void SetServerStatus(string status)
        {
            SqlCommand command = new SqlCommand(Program.DB.setServerStatus);
            command.Parameters.AddWithValue("@status", status);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            Program.DB.executeNonQuery(command);
        }

        public void SetServerSoftwareVersion(string version)
        {
            SqlCommand command = new SqlCommand(Program.DB.setServerSoftwareVersion);
            command.Parameters.AddWithValue("@version", version);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            Program.DB.executeNonQuery(command);
        }

        public void SetServerLastWindowsUpdate(DateTime? lastUpdateSuccess)
        {
            SqlCommand command = new SqlCommand(Program.DB.updateServer);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            command.Parameters.AddWithValue("@last_windows_update", lastUpdateSuccess);
            Program.DB.executeNonQuery(command);
        }

        public string GetServerStatus()
        {
            SqlCommand command = new SqlCommand(Program.DB.queryGetServerStatus);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return "";
            }

            DataRow row = rows.Rows[0];

            return row["status"].ToString();
        }

        public List<string> GetExemptServices()
        {
            List<string> svcs = new List<string>();

            try
            {
                SqlCommand command = new SqlCommand(Program.DB.queryGetExemptServices);
                command.Parameters.AddWithValue("@hostname", Environment.MachineName);
                DataTable rows = Program.DB.executeQuery(command);
                
                if (rows.Rows.Count < 1)
                {
                    return svcs;
                }

                foreach (DataRow row in rows.Rows)
                {
                    svcs.Add(row["name"].ToString());
                }

                return svcs;
            }

            catch 
            {
                return svcs;
            }

            
        }


        private string GetIPAddress()
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

        public int getServerId()
        {
            int id;
            SqlCommand command = new SqlCommand(Program.DB.queryGetServerId);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return 0;
            }

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        public int getServerDiskId(DriveInfo drive)
        {
            int id;
            SqlCommand command = new SqlCommand(Program.DB.queryGetServerDiskId);
            command.Parameters.AddWithValue("@server_id", getServerId());
            command.Parameters.AddWithValue("@name", drive.Name);

            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return 0;
            }

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        public int getServerServiceId(ServiceControllerEx svc)
        {
            int id;
            SqlCommand command = new SqlCommand(Program.DB.queryGetServerServiceId);
            command.Parameters.AddWithValue("@server_id", getServerId());
            command.Parameters.AddWithValue("@name", svc.DisplayName);

            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return 0;
            }

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        public int getUpdateId(IUpdate update)
        {
            int id;
            SqlCommand command = new SqlCommand(Program.DB.queryGetUpdateId);
            command.Parameters.AddWithValue("@title", update.Title);
            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return 0;
            }
            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        public void storeServerDiskInDashboard(DriveInfo drive)
        {
            if ( DriveExistsInDashboard(drive) )
            {
                updateServerDiskInDashboard(drive);
                return;
            }

            SqlCommand command = new SqlCommand(Program.DB.insertServerDisk);
            command.Parameters.AddWithValue("@name", drive.Name);
            command.Parameters.AddWithValue("@label", drive.VolumeLabel);
            command.Parameters.AddWithValue("@server_id", getServerId() );
            command.Parameters.AddWithValue("@size_gb", drive.TotalSize / 1073741824 );
            command.Parameters.AddWithValue("@used_gb", ( drive.TotalSize - drive.TotalFreeSpace ) / 1073741824 );
            command.Parameters.AddWithValue("@free_gb", ( drive.TotalFreeSpace ) / 1073741824 );
            Program.DB.executeNonQuery(command);
      
        }

        public void updateServerDiskInDashboard(DriveInfo drive)
        {
            int id = getServerDiskId(drive);

            SqlCommand command = new SqlCommand(Program.DB.updateServerDisk);
            command.Parameters.AddWithValue("@label", drive.VolumeLabel);
            command.Parameters.AddWithValue("@server_id", getServerId());
            command.Parameters.AddWithValue("@size_gb", drive.TotalSize / 1073741824);
            command.Parameters.AddWithValue("@used_gb", (drive.TotalSize - drive.TotalFreeSpace) / 1073741824);
            command.Parameters.AddWithValue("@free_gb", (drive.TotalFreeSpace) / 1073741824);
            command.Parameters.AddWithValue("@id", id);

            Program.DB.executeNonQuery(command);
        }

        public void storeUpdateInDashboard(IUpdate update)
        {
            if (!this.isUpdateInDashboard(update))
            {
                SqlCommand command = new SqlCommand(Program.DB.insertUpdate);
                command.Parameters.AddWithValue("@title", update.Title);
                command.Parameters.AddWithValue("@description", update.Title);
                command.Parameters.AddWithValue("@kb_article", this.getKBArticle(update.Title));

                Program.DB.executeNonQuery(command);
            }
        }

        private bool DriveExistsInDashboard(DriveInfo drive)
        {
            return getServerDiskId(drive) > 0;
        }

        public void setUpdateInstalledAt(IUpdate update, int install_batch)
        {
            SqlCommand command = new SqlCommand(Program.DB.setUpdateInstalledAt);
            command.Parameters.AddWithValue("@install_batch", install_batch);
            command.Parameters.AddWithValue("@update_id", this.getUpdateId(update));
            command.Parameters.AddWithValue("@server_id", this.getServerId());
            Program.DB.executeNonQuery(command);
        }

        public dynamic storeUpdateInDashboardForThisServer(IUpdate update)
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

            SqlCommand command = new SqlCommand(Program.DB.insertUpdateForServer);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            command.Parameters.AddWithValue("@eula_accepted", update.EulaAccepted);
            command.Parameters.AddWithValue("@downloaded_flag", update.IsDownloaded);
            command.Parameters.AddWithValue("@hidden_flag", update.IsHidden);
            command.Parameters.AddWithValue("@installed_flag", update.IsInstalled);
            command.Parameters.AddWithValue("@mandatory_flag", update.IsMandatory);
            command.Parameters.AddWithValue("@max_download_size", update.MaxDownloadSize);
            command.Parameters.AddWithValue("@min_download_size", update.MinDownloadSize);

            Program.DB.executeNonQuery(command);

            return true;
        }

        public bool isUpdateInDashboard(IUpdate update)
        {
            int count;
            SqlCommand command = new SqlCommand(Program.DB.countUpdates);
            command.Parameters.AddWithValue("@title", update.Title);
            DataTable rows = Program.DB.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        public int getLastInstallBatch()
        {
            int batch;                     

            SqlCommand command = new SqlCommand(Program.DB.queryGetLastInstallBatch);
            command.Parameters.AddWithValue("@server_id", this.getServerId());
            DataTable rows = Program.DB.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["install_batch"].ToString(), out batch);

            return (batch > 0) ? batch : 0;
        }

        public bool isUpdateInDashboardForThisServer(IUpdate update)
        {
            int count;
            SqlCommand command = new SqlCommand(Program.DB.countUpdatesForServer);
            command.Parameters.AddWithValue("@title", update.Title);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = Program.DB.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        public bool isUpdateApproved(IUpdate update)
        {
            int count;

            int update_id = this.getUpdateId(update);
            int server_id = this.getServerId();

            SqlCommand command = new SqlCommand(Program.DB.queryUpdateApproved);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            DataTable rows = Program.DB.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        public bool isUpdateHidden(IUpdate update)
        {
            int count;

            int update_id = this.getUpdateId(update);
            int server_id = this.getServerId();

            SqlCommand command = new SqlCommand(Program.DB.queryUpdateHidden);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            DataTable rows = Program.DB.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        public dynamic updateUpdateInDashboardForThisServer(IUpdate update)
        {
            if (!this.isUpdateInDashboardForThisServer(update))
            {
                return this.storeUpdateInDashboardForThisServer(update);
            }

            // first get update_id
            int update_id = this.getUpdateId(update);

            // get server_id
            int server_id = this.getServerId();


            SqlCommand command = new SqlCommand(Program.DB.updateUpdateForServer);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            command.Parameters.AddWithValue("@eula_accepted", update.EulaAccepted);
            command.Parameters.AddWithValue("@downloaded_flag", update.IsDownloaded);
            command.Parameters.AddWithValue("@hidden_flag", update.IsHidden);
            command.Parameters.AddWithValue("@installed_flag", update.IsInstalled);
            command.Parameters.AddWithValue("@mandatory_flag", update.IsMandatory);
            command.Parameters.AddWithValue("@max_download_size", update.MaxDownloadSize);
            command.Parameters.AddWithValue("@min_download_size", update.MinDownloadSize);

            Program.DB.executeNonQuery(command);

            return true;
        }

        public string getKBArticle(string title)
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
    }
}
