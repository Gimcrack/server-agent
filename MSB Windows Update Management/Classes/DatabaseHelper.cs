using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;

namespace MSB_Windows_Update_Management
{
    public partial class DatabaseHelper
    {
        public string connectionString;
        public string countUpdates;
        public string countUpdatesForServer;
        public string queryGetUpdateId;
        public string queryGetServerId;
        public string queryGetServerStatus;
        public string queryGetServerDiskId;
        public string queryGetServerServiceId;
        public string queryUpdateApproved;
        public string queryUpdateHidden;
        public string queryGetEmailNotifications;
        public string queryGetTextNotifications;
        public string queryGetExemptServices;
        public string queryGetLastInstallBatch;

        public string insertUpdate;
        public string insertUpdateForServer;
        public string insertServer;
        public string insertUpdateBatch;
        public string insertAlert;
        public string insertServerDisk;
        public string updateServerDisk;
        public string updateServer;
        public string updateUpdateForServer;
        public string setServerStatus;
        public string setServerAlert;
        public string setServerSoftwareVersion;
        public string setUpdateInstalledAt;
        public string deleteSupersededUpdates;
        public string checkIn;

        public DatabaseHelper()
        {
            connectionString = "Server=msbsqlrpt;Database=ITDashboard;User Id=ITDashboardUser;Password=ITDashboardUser;";

            //connectionString = "Server=msbsqltst;Database=ITDashboardDev;User Id=ITDashboardDevUser;Password=ITDashboardDevUser;";

            countUpdates = "SELECT count(*) as count FROM [dbo].[updates] WHERE title = @title";

            countUpdatesForServer = "SELECT count(*) as count " +
            "FROM [dbo].[update_details] as ud " +
            "INNER JOIN dbo.updates as u " +
            "ON u.id = ud.update_id " +
            "INNER JOIN dbo.servers as s " +
            "ON s.id = ud.server_id " +
            "WHERE u.title = @title " +
            "AND s.name = @hostname";

            queryGetLastInstallBatch = "SELECT isnull(max(install_batch),0) as install_batch FROM update_details WHERE server_id=@server_id";

            queryGetUpdateId = "SELECT id FROM dbo.updates WHERE title = @title";

            queryGetServerId = "SELECT id FROM dbo.servers WHERE name = @hostname";

            queryGetServerStatus = "SELECT status FROM dbo.servers WHERE name = @hostname";

            queryGetServerDiskId = "SELECT id FROM dbo.server_disks WHERE server_id = @server_id AND name = @name";

            queryGetServerServiceId = "SELECT id FROM dbo.server_services WHERE server_id = @server_id AND name = @name";

            queryUpdateApproved = "SELECT count(id) as count FROM dbo.update_details WHERE installed_flag=0 AND approved_flag=1 AND server_id=@server_id AND update_id=@update_id";

            queryUpdateHidden = "SELECT count(id) as count FROM dbo.update_details WHERE hidden_flag=1 AND server_id=@server_id AND update_id=@update_id";

            queryGetEmailNotifications = "SELECT email FROM dbo.notifications WHERE notifications_enabled in ('Both','Email')";

            queryGetTextNotifications = "SELECT phone_number FROM dbo.notifications WHERE notifications_enabled in ('Both','Text')";

            queryGetExemptServices = "SELECT name FROM dbo.notification_exemptions WHERE server_name in (@hostname,'All')";

            insertUpdate = "INSERT INTO [dbo].[updates] (title, description, kb_article, created_at, updated_at) VALUES (@title, @description, @kb_article, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

            insertUpdateForServer = "INSERT INTO [dbo].[update_details] (update_id, server_id, eula_accepted, downloaded_flag, hidden_flag, installed_flag, mandatory_flag, max_download_size, min_download_size, approved_flag, created_at, updated_at) VALUES ( " +
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

            insertServer = "INSERT INTO dbo.servers (name, ip, operating_system_id, created_at, updated_at) VALUES (@hostname, @ip, 1, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

            insertUpdateBatch = "INSERT INTO dbo.update_batches (server_id, batch_id, result, created_at, updated_at) VALUES (@server_id, @batch_id, @result, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

            insertAlert = "INSERT INTO dbo.alerts (message, alertable_type, alertable_id, created_at, updated_at) VALUES (@message, @alertable_type, @alertable_id, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

            insertServerDisk = "INSERT INTO dbo.server_disks ( name, label, server_id, size_gb, used_gb, free_gb, created_at, updated_at) VALUES (@name, @label, @server_id, @size_gb, @used_gb, @free_gb, CAST(SYSDATETIME() AS VARCHAR(19)), CAST(SYSDATETIME() AS VARCHAR(19)))";

            updateServerDisk = "UPDATE dbo.server_disks SET label=@label, size_gb=@size_gb, used_gb=@used_gb, free_gb=@free_gb, updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) where id=@id";

            updateServer = "UPDATE dbo.servers SET " +
                "last_windows_update = @last_windows_update, " +
                "updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) " +
                "WHERE name = @hostname";

            updateUpdateForServer = "UPDATE [dbo].[update_details] SET " +
                "eula_accepted = @eula_accepted, " +
                "downloaded_flag = @downloaded_flag, " +
                "hidden_flag = @hidden_flag, " +
                "installed_flag = @installed_flag, " +
                "mandatory_flag = @mandatory_flag, " +
                "updated_at = CAST(SYSDATETIME() AS VARCHAR(19)) " +
                "WHERE server_id = @server_id AND update_id = @update_id and superseded_flag = 0";

            setServerStatus = "UPDATE dbo.servers SET status=@status, updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) WHERE name = @hostname";

            checkIn = "UPDATE dbo.servers SET updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) WHERE name = @hostname";

            setServerAlert = "UPDATE dbo.servers SET alert=@alert, updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) WHERE name = @hostname";

            setServerSoftwareVersion = "UPDATE dbo.servers SET software_version=@version, updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) WHERE name = @hostname";

            setUpdateInstalledAt = "UPDATE update_details SET install_batch=@install_batch, installed_at = CAST(SYSDATETIME() AS VARCHAR(19)) WHERE server_id = @server_id AND update_id = @update_id";

            deleteSupersededUpdates = "UPDATE dbo.update_details " +
                                      "SET hidden_flag = 1, superseded_flag = 1 " + 
                                      "WHERE server_id = @server_id AND " + 
                                      "update_id NOT IN ({0}) AND " +
                                      "installed_flag = 0 AND approved_flag = 1";

        }

        public void executeNonQuery(SqlCommand command)
        {
            using (SqlConnection dbc = new SqlConnection(this.connectionString))
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
                    Program.Events.WriteEntry(e.Message, EventLogEntryType.Error);
                }

                finally
                {
                    dbc.Close();
                }
            }
        }

        public DataTable executeQuery(SqlCommand command)
        {
            using (SqlConnection dbc = new SqlConnection(this.connectionString))
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

                    catch (SqlException e)
                    {
                        Program.Events.WriteEntry(e.Message, EventLogEntryType.Error);
                        reader.Close();
                        return new DataTable();
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

    }
}
