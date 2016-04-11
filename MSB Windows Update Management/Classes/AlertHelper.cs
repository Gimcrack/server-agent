using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;


namespace MSB_Windows_Update_Management
{
    public partial class AlertHelper
    {

        public void ServerAlert(string errorMessage)
        {            
            if ( ! Fire(errorMessage, @"App\Server", Program.Dash.getServerId() ) )
            {
                AlertFallback(errorMessage);
            }
        }

        public void DatabaseAlert(string errorMessage, int databaseId)
        {
            if ( ! Fire(errorMessage, @"App\Database", databaseId) )
            {
                AlertFallback(errorMessage);
            }   
        }

        public bool Fire(string errorMessage, string alertableType, int alertableId)
        {
            try
            {
                SqlCommand command = new SqlCommand(Program.DB.insertAlert);
                command.Parameters.AddWithValue("@message", errorMessage);
                command.Parameters.AddWithValue("@alertable_type", alertableType);
                command.Parameters.AddWithValue("@alertable_id", alertableId);
                Program.DB.executeNonQuery(command);
                return Program.IsDashboardOnline();
            }

            catch( Exception e )
            {
                return false;
            }
            
        }

        public void AlertFallback( string errorMessage )
        {
            Program.Mail.alert(errorMessage);
            Program.Msg.alert(errorMessage);
        }

        
    }
}
