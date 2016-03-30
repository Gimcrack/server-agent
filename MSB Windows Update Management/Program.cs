using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace MSB_Windows_Update_Management
{
    static class Program
    {
        public static MailHelper Mail = new MailHelper();
        public static MessageHelper Msg = new MessageHelper();
        public static UpdateHelper Upd = new UpdateHelper();
        public static DashboardHelper Dash = new DashboardHelper();
        public static DatabaseHelper DB = new DatabaseHelper();
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
    }
}
