using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WUApiLib;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace MSB_Windows_Update_Management
{
    public partial class UpdateHelper
    {
        public DatabaseHelper db = new DatabaseHelper();
        public DashboardHelper dashb = new DashboardHelper();
        private System.Diagnostics.EventLog eventLog1;

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

        public UpdateHelper()
        {
            // set up the event log
            eventLog1 = new System.Diagnostics.EventLog();
            eventLog1.Source = "Updates";
            eventLog1.Log = "MSB";
        }

        public void LookForUpdates()
        {
            ShouldInstallUpdates = (dashb.GetServerStatus() == "Ready For Updates");

            uDownloadList = new UpdateCollection();
            uInstallList = new UpdateCollection();
            uSearcher = uSession.CreateUpdateSearcher();
            List<String> updateList = new List<string>();


            dashb.status("Looking For Available Updates...");

            uSearcher.Online = true;
            uResult = uSearcher.Search("IsInstalled=0 and Type='Software'");

            dashb.status("There are " + uResult.Updates.Count + " available updates");


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
            dashb.UpdateServerInfo();

            eventLog1.WriteEntry(string.Join(System.Environment.NewLine, updateList), EventLogEntryType.Information);
            Console.WriteLine(string.Join(System.Environment.NewLine, updateList));

            if (uSysInfo.RebootRequired)
            {
                dashb.status("Reboot Required");
            }
            else
            {
                dashb.status("Idle");
            }

        }

        private void processUpdate(IUpdate update, List<String> updateList, UpdateCollection uDownloadList, UpdateCollection uInstallList)
        {

            dashb.storeUpdateInDashboard(update);

            dashb.updateUpdateInDashboardForThisServer(update);

            if (!update.IsDownloaded)
            {
                uDownloadList.Add(update);
            }

            if (dashb.isUpdateApproved(update))
            {
                uInstallList.Add(update);
            }

            updateList.Add(update.Title.ToString());
        }

        private void downloadUpdates(UpdateCollection updateList)
        {
            try
            {
                dashb.status("Downloading " + updateList.Count + " Updates");
                uDownloader = uSession.CreateUpdateDownloader();
                uDownloader.Updates = updateList;
                uDownloader.Download();
            }

            catch (COMException e)
            {
                eventLog1.WriteEntry(e.Message, EventLogEntryType.Error);
            }
        }

        private void updateUpdates()
        {
            // TODO: Insert monitoring activities here.
            dashb.status("Updating updates...");

            uSearcher = uSession.CreateUpdateSearcher();
            uSearcher.Online = true;
            uResult = uSearcher.Search("Type='Software'");

            foreach (IUpdate update in uResult.Updates)
            {
                update.IsHidden = dashb.isUpdateHidden(update);

                if (dashb.isUpdateInDashboardForThisServer(update))
                {
                    eventLog1.WriteEntry("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title, EventLogEntryType.Information);
                    Console.WriteLine("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title);
                    dashb.updateUpdateInDashboardForThisServer(update);
                }
            }
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

        

        private void installUpdates(UpdateCollection updateList)
        {


            if ((updateList.Count > 0) && ShouldInstallUpdates)
            {
                dashb.status("Installing Updates...");

                try
                {
                    dashb.status("Installing " + updateList.Count + " Updates");
                    uInstaller = uSession.CreateUpdateInstaller();
                    uInstaller.AllowSourcePrompts = false;
                    if (uInstaller.IsBusy)
                    {
                        dashb.status("Update Installer Busy, Try Again");
                        return;
                    }

                    if (uInstaller.RebootRequiredBeforeInstallation)
                    {
                        dashb.status("Reboot Required");
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

        

        
    }
}
