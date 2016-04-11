using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Net;

namespace MSB_Windows_Update_Management
{
    static class Program
    {
        public static string DashboardAddress = "http://itdashboard/api/v1/Server";
        private static string Username = @"svckbox";
        private static string Password = @"D3nv3r0melet2050!!";



        public static AlertHelper Alert = new AlertHelper();
        public static MailHelper Mail = new MailHelper();
        public static MessageHelper Msg = new MessageHelper();
        public static UpdateHelper Upd = new UpdateHelper();
        public static DashboardHelper Dash = new DashboardHelper();
        public static DatabaseHelper DB = new DatabaseHelper();
        public static DiskHelper Disk = new DiskHelper();
        public static ServiceHelper Serv = new ServiceHelper();
        public static EventLog Events = new EventLog();
        

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            SetupEventLog();

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new MSBWindowsUpdateManagement() 
            };
            ServiceBase.Run(ServicesToRun);
        }

        /// <summary>
        /// Setup the event log for the application
        /// </summary>
        static void SetupEventLog()
        {
            if (!EventLog.SourceExists("Updates"))
            {
                EventLog.CreateEventSource("Updates", "MSB");
            }
            Program.Events.Source = "Updates";
            Program.Events.Log = "MSB";
        }

        /// <summary>
        /// Is the dashboard online?
        /// </summary>
        /// <returns></returns>
        public static bool IsDashboardOnline()
        {
            return RemoteFileExists(Program.DashboardAddress);
        }

        ///
        /// Checks the file exists or not.
        ///
        /// The URL of the remote file.
        /// True : If the file exits, False if file not exists
        private static bool RemoteFileExists(string url)
        {
            try
            {
                //Creating the HttpWebRequest
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                CredentialCache creds = new CredentialCache();
                creds.Add(new System.Uri(url), "Basic", new System.Net.NetworkCredential(Username, Password, "msb"));

                //request.PreAuthenticate = true;
                request.Credentials = creds; 

                //Setting the Request method HEAD, you can also use GET too.
                request.Method = "HEAD";
                //Getting the Web Response.
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                //Returns TRUE if the Status code == 200
                response.Close();
                Console.WriteLine(response.StatusCode);
                Console.WriteLine(response.StatusDescription);
                return (response.StatusCode == HttpStatusCode.OK);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                //Any exception will returns false.
                return false;
            }
        }
    }
}
