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
        /// <summary>
        /// MSB IT Windows Update Management Service
        /// 
        /// Automatically: 
        ///  -- Checks for updates every 8 hours
        ///  -- Downloads missing updates
        ///  -- Logs all available updates to the IT Dashboard database
        ///  -- Installs approved updates on designated servers
        ///  -- Reboots designated servers at the designated time
        ///  
        /// To Do:
        ///  -- Add service monitoring
        /// </summary>
        public MSBWindowsUpdateManagement()
        {
            InitializeComponent();
            SetupEventLog();

            // The service executable can also be run
            //  as a console application
            if (Environment.UserInteractive)
            {
                upd.LookForUpdates();
            }
        }
        
        /// <summary>
        /// Most Windows Update functions are handled
        ///  by the UpdateHelper class
        /// </summary>
        public UpdateHelper upd = new UpdateHelper();
                
        /// <summary>
        /// The service logs everything to a custom event log
        ///  Log = MSB, Source = Updates
        /// </summary>
        private System.Diagnostics.EventLog eventLog1;
        private void SetupEventLog()
        {
            if (!System.Diagnostics.EventLog.SourceExists("Updates"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "Updates", "MSB");
            }
            eventLog1 = new System.Diagnostics.EventLog();
            eventLog1.Source = "Updates";
            eventLog1.Log = "MSB";
        }

        private System.ComponentModel.IContainer components = null;

        

        
        /// <summary>
        /// This method automatically fires whenever the
        ///  service is started
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("Started Service MSB Windows Update Management");
            this.SetupTimers();
            
            
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);


            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

        }

        /// <summary>
        /// There are 3 main timers that control
        ///  when the service performs work
        ///  
        /// lookForUpdates - Controls when the service looks for updates
        ///  -- initial run, 12 sec after service start
        ///  -- following runs, every 8 hours
        ///  
        ///  updateServerInfo - Controls when the service updates server status info
        ///  -- initial run, 5 sec after service start
        ///  -- following runs, every 10 min
        ///  
        ///  getServerStatus - Controls when the service checks the server status
        ///  -- initial run, 2 sec after service start
        ///  -- following runs, every 2 min
        ///  
        /// </summary>
        public System.Timers.Timer lookForUpdatesTimer = new System.Timers.Timer();
        public System.Timers.Timer updateServerInfoTimer = new System.Timers.Timer();
        public System.Timers.Timer getServerStatusTimer = new System.Timers.Timer();
        private void SetupTimers()
        {
            // Check for available updates timer    
            lookForUpdatesTimer.Interval = 12000; 
            lookForUpdatesTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnLookForUpdatesTimer);
            lookForUpdatesTimer.Start();

            // Update Server Info timer    
            updateServerInfoTimer.Interval = 5000; 
            updateServerInfoTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnUpdateServerInfoTimer);
            updateServerInfoTimer.Start();

            // Update Server Info timer    
            getServerStatusTimer.Interval = 1000; 
            getServerStatusTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnGetServerStatusTimer);
            getServerStatusTimer.Start();
        }

        /// <summary>
        /// On Look For Updates Timer
        /// -- instructs the service to look for updates
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void OnLookForUpdatesTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            lookForUpdatesTimer.Interval = 1000 * 60 * 60 * 8; // 8 hours
            upd.LookForUpdates();    
        }

        /// <summary>
        /// On Update Server Info
        /// -- instructs the service to update server info
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void OnUpdateServerInfoTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            updateServerInfoTimer.Interval = 1000 * 60 * 10; // 10 min
            upd.dashb.UpdateServerInfo();
        }

        /// <summary>
        /// On Get Server Status
        /// -- instructs the service to get the server status
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void OnGetServerStatusTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            getServerStatusTimer.Interval = 1000 * 60 * 2; // 2 mins
            
            switch ( upd.dashb.GetServerStatus() )
            {
                case "Ready For Updates" :
                case "Look For Updates" :
                    upd.LookForUpdates();
                    break;

                case "Ready For Reboot" :
                    upd.dashb.RebootComputer();
                    break;
            }    
        }


        /**
         *  Boilerplate below here
         */

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
