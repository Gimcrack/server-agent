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
        public string queryUpdateApproved;
        public string queryUpdateHidden;

        public string insertUpdate;
        public string insertUpdateForServer;
        public string insertServer;
        public string updateServer;
        public string updateUpdateForServer;
        public string setServerStatus;

        private System.Diagnostics.EventLog eventLog1;
        

        public DatabaseHelper()
        {
            // set up the event log
            eventLog1 = new System.Diagnostics.EventLog();
            eventLog1.Source = "Updates";
            eventLog1.Log = "MSB";
            
            connectionString = "Server=msbsqlrpt;Database=ITDashboard;User Id=ITDashboardUser;Password=ITDashboardUser;";

            countUpdates = "SELECT count(*) as count FROM [dbo].[updates] WHERE title = @title";

            countUpdatesForServer = "SELECT count(*) as count " +
            "FROM [dbo].[update_details] as ud " +
            "INNER JOIN dbo.updates as u " +
            "ON u.id = ud.update_id " +
            "INNER JOIN dbo.servers as s " +
            "ON s.id = ud.server_id " +
            "WHERE u.title = @title " +
            "AND s.name = @hostname";

            queryGetUpdateId = "SELECT id FROM dbo.updates WHERE title = @title";

            queryGetServerId = "SELECT id FROM dbo.servers WHERE name = @hostname";

            queryGetServerStatus = "SELECT status FROM dbo.servers WHERE name = @hostname";

            queryUpdateApproved = "SELECT count(id) as count FROM dbo.update_details WHERE installed_flag=0 AND approved_flag=1 AND server_id=@server_id AND update_id=@update_id";

            queryUpdateHidden = "SELECT count(id) as count FROM dbo.update_details WHERE hidden_flag=1 AND server_id=@server_id AND update_id=@update_id";

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
                "WHERE server_id = @server_id AND update_id = @update_id";

            setServerStatus = "UPDATE dbo.servers SET status=@status, updated_at=CAST(SYSDATETIME() AS VARCHAR(19)) WHERE name = @hostname";

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
                    eventLog1.WriteEntry(e.Message, EventLogEntryType.Error);
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
