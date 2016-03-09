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
using WUApiLib;

namespace MSB_Windows_Update_Management
{
    public partial class DashboardHelper
    {
        public DatabaseHelper db = new DatabaseHelper();
        private System.Diagnostics.EventLog eventLog1;
        public AutomaticUpdatesClass uAutomaticUpdate = new AutomaticUpdatesClass();


        public DashboardHelper()
        {
            // set up the event log
            eventLog1 = new System.Diagnostics.EventLog();
            eventLog1.Source = "Updates";
            eventLog1.Log = "MSB";
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

        public void status(string message, EventLogEntryType type = EventLogEntryType.Information)
        {
            eventLog1.WriteEntry(message, type);
            Console.WriteLine(message);
            this.SetServerStatus(message);
        }

        public void storeServerInDashboard()
        {
            SqlCommand command = new SqlCommand(db.insertServer);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            command.Parameters.AddWithValue("@ip", this.GetIPAddress());
            db.executeNonQuery(command);
        }

        public void SetServerStatus(string status)
        {
            SqlCommand command = new SqlCommand(db.setServerStatus);
            command.Parameters.AddWithValue("@status", status);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            db.executeNonQuery(command);
        }

        public void SetServerLastWindowsUpdate(DateTime? lastUpdateSuccess)
        {
            SqlCommand command = new SqlCommand(db.updateServer);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            command.Parameters.AddWithValue("@last_windows_update", lastUpdateSuccess);
            db.executeNonQuery(command);
        }

        public string GetServerStatus()
        {
            SqlCommand command = new SqlCommand(db.queryGetServerStatus);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = db.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return "";
            }

            DataRow row = rows.Rows[0];

            return row["status"].ToString();
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
            SqlCommand command = new SqlCommand(db.queryGetServerId);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = db.executeQuery(command);

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
            SqlCommand command = new SqlCommand(db.queryGetUpdateId);
            command.Parameters.AddWithValue("@title", update.Title);
            DataTable rows = db.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return 0;
            }
            DataRow row = rows.Rows[0];

            Int32.TryParse(row["id"].ToString(), out id);

            if (id > 0) return id;
            return 0;
        }

        public void storeUpdateInDashboard(IUpdate update)
        {
            if (!this.isUpdateInDashboard(update))
            {
                SqlCommand command = new SqlCommand(db.insertUpdate);
                command.Parameters.AddWithValue("@title", update.Title);
                command.Parameters.AddWithValue("@description", update.Title);
                command.Parameters.AddWithValue("@kb_article", this.getKBArticle(update.Title));

                db.executeNonQuery(command);
            }
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

            SqlCommand command = new SqlCommand(db.insertUpdateForServer);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            command.Parameters.AddWithValue("@eula_accepted", update.EulaAccepted);
            command.Parameters.AddWithValue("@downloaded_flag", update.IsDownloaded);
            command.Parameters.AddWithValue("@hidden_flag", update.IsHidden);
            command.Parameters.AddWithValue("@installed_flag", update.IsInstalled);
            command.Parameters.AddWithValue("@mandatory_flag", update.IsMandatory);
            command.Parameters.AddWithValue("@max_download_size", update.MaxDownloadSize);
            command.Parameters.AddWithValue("@min_download_size", update.MinDownloadSize);

            db.executeNonQuery(command);

            return true;
        }

        public bool isUpdateInDashboard(IUpdate update)
        {
            int count;
            SqlCommand command = new SqlCommand(db.countUpdates);
            command.Parameters.AddWithValue("@title", update.Title);
            DataTable rows = db.executeQuery(command);

            DataRow row = rows.Rows[0];

            Int32.TryParse(row["count"].ToString(), out count);

            if (count > 0) return true;
            return false;
        }

        public bool isUpdateInDashboardForThisServer(IUpdate update)
        {
            int count;
            SqlCommand command = new SqlCommand(db.countUpdatesForServer);
            command.Parameters.AddWithValue("@title", update.Title);
            command.Parameters.AddWithValue("@hostname", Environment.MachineName);
            DataTable rows = db.executeQuery(command);

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

            SqlCommand command = new SqlCommand(db.queryUpdateApproved);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            DataTable rows = db.executeQuery(command);

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

            SqlCommand command = new SqlCommand(db.queryUpdateHidden);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            DataTable rows = db.executeQuery(command);

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


            SqlCommand command = new SqlCommand(db.updateUpdateForServer);
            command.Parameters.AddWithValue("@update_id", update_id);
            command.Parameters.AddWithValue("@server_id", server_id);
            command.Parameters.AddWithValue("@eula_accepted", update.EulaAccepted);
            command.Parameters.AddWithValue("@downloaded_flag", update.IsDownloaded);
            command.Parameters.AddWithValue("@hidden_flag", update.IsHidden);
            command.Parameters.AddWithValue("@installed_flag", update.IsInstalled);
            command.Parameters.AddWithValue("@mandatory_flag", update.IsMandatory);
            command.Parameters.AddWithValue("@max_download_size", update.MaxDownloadSize);
            command.Parameters.AddWithValue("@min_download_size", update.MinDownloadSize);

            db.executeNonQuery(command);

            return true;
        }

        public void RebootComputer()
        {
            this.status("Rebooting in 5 min");
            System.Diagnostics.Process.Start("shutdown.exe", "-r -t 300 -c \"Rebooting in 5. Run 'shutdown -a' to abort.\"");
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
