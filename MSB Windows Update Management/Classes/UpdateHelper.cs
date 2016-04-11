using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WUApiLib;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Mail;

namespace MSB_Windows_Update_Management
{
    public partial class UpdateHelper
    {
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
        }

        public void LookForUpdates()
        {
            // make sure the server is supposed to get updates
            if ( File.Exists(@"c:\scripts\MSBUpdateManager\_no_updates.txt"))
            {
                return;
            }
            
            ShouldInstallUpdates = (Program.Dash.GetServerStatus() == "Ready For Updates");

            uDownloadList = new UpdateCollection();
            uInstallList = new UpdateCollection();
            uSearcher = uSession.CreateUpdateSearcher();
            List<String> updateList = new List<string>();
            List<int> update_ids = new List<int>();

            update_ids.Add(-1);


            Program.Dash.status("Looking For Available Updates...");

            uSearcher.Online = true;
            uResult = uSearcher.Search("IsInstalled=0 and Type='Software'");

            if (uResult.Updates.Count > 0)
            {
                Program.Dash.status("There are " + uResult.Updates.Count + " available updates");

                updateList.Add("Available Updates...");
                foreach (IUpdate update in uResult.Updates)
                {
                    this.processUpdate(update, updateList, uDownloadList, uInstallList);
                    update_ids.Add(Program.Dash.getUpdateId(update));
                }

                Program.Dash.status("Cleaning up superseded updates.");
                Program.Dash.CleanupUpdates(update_ids);

                // download any missing updates
                this.downloadUpdates(uDownloadList);

                // install any queued updates
                this.installUpdates(uInstallList);

                // update the status of updates
                this.UpdateUpdates(uResult.Updates);

                Program.Events.WriteEntry(string.Join(System.Environment.NewLine, updateList), EventLogEntryType.Information);
                Console.WriteLine(string.Join(System.Environment.NewLine, updateList));
            }

            Program.Dash.status("Cleaning up superseded updates.");
            Program.Dash.CleanupUpdates(update_ids);

            // update server info
            Program.Dash.UpdateServerInfo();
            this.SetServerStatusNominal();

        }


        public void SetServerStatusNominal()
        {
            if (uSysInfo.RebootRequired)
            {
                Program.Dash.status("Reboot Required");
            }
            else
            {
                Program.Dash.status("Idle");
            }
        }

        private void processUpdate(IUpdate update, List<String> updateList, UpdateCollection uDownloadList, UpdateCollection uInstallList)
        {

            Program.Dash.storeUpdateInDashboard(update);

            Program.Dash.updateUpdateInDashboardForThisServer(update);

            if (!update.IsDownloaded)
            {
                uDownloadList.Add(update);
            }

            if (Program.Dash.isUpdateApproved(update))
            {
                uInstallList.Add(update);
            }

            updateList.Add(update.Title.ToString());
        }

        private void downloadUpdates(UpdateCollection updateList)
        {
            try
            {
                Program.Dash.status("Downloading " + updateList.Count + " Updates");
                uDownloader = uSession.CreateUpdateDownloader();
                uDownloader.Updates = updateList;
                uDownloader.Download();
            }

            catch (COMException e)
            {
                Program.Events.WriteEntry(e.Message, EventLogEntryType.Error);
            }
        }

        public void UpdateUpdates(UpdateCollection updates)
        {
            foreach (IUpdate update in updates)
            {
                update.IsHidden = Program.Dash.isUpdateHidden(update);

                if (Program.Dash.isUpdateInDashboardForThisServer(update))
                {
                    Program.Events.WriteEntry("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title, EventLogEntryType.Information);
                    Console.WriteLine("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title);
                    Program.Dash.updateUpdateInDashboardForThisServer(update);
                }
            }
        }

        public void UpdateAllUpdates()
        {
            // TODO: Insert monitoring activities here.
            Program.Dash.status("Updating updates...");

            uSearcher = uSession.CreateUpdateSearcher();
            uSearcher.Online = true;
            uResult = uSearcher.Search("Type='Software'");

            foreach (IUpdate update in uResult.Updates)
            {
                update.IsHidden = Program.Dash.isUpdateHidden(update);

                if (Program.Dash.isUpdateInDashboardForThisServer(update))
                {
                    Program.Events.WriteEntry("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title, EventLogEntryType.Information);
                    Console.WriteLine("Updating update in the dashboard for server: " + Environment.MachineName + ", update : " + update.Title);
                    Program.Dash.updateUpdateInDashboardForThisServer(update);
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
            uInstaller = uSession.CreateUpdateInstaller();
            uInstaller.AllowSourcePrompts = false;

            if ( ! ShouldInstallUpdates || updateList.Count < 1 )
            {
                return;
            }

            if (uInstaller.IsBusy)
            {
                Program.Dash.status("Update Installer Busy, Try Again");
                return;
            }

            if (uInstaller.RebootRequiredBeforeInstallation)
            {
                Program.Dash.status("Reboot Required");
                return;
            }

            Program.Dash.status("Installing " + updateList.Count + " Updates");
     
            int counter = 0;
            int installBatch = Program.Dash.getLastInstallBatch() + 1;
            List<string> results = new List<string>();

            foreach( IUpdate update in updateList )
            {
                counter++;
                string resultText;

                Program.Dash.status(string.Format("Installing update {0} of {1} : {2}", counter.ToString(), updateList.Count.ToString(), update.Title));
                resultText = this.install(update, installBatch);
                results.Add( string.Format("Result for '{0}' : {1}",update.Title,resultText) );
            }

            // send the result notification
            string resultMessage = this.getResultMessage(results);
            Program.Dash.status(resultMessage);
            Program.Dash.InsertUpdateBatch(resultMessage, installBatch);
        }

        private string install(IUpdate update, int installBatch)
        {
            string resultText;

            try
            {
                UpdateCollection updates = new UpdateCollection();
                updates.Add(update);
                uInstaller = uSession.CreateUpdateInstaller();
                uInstaller.Updates = updates;
                IInstallationResult result = uInstaller.Install();
                IUpdateInstallationResult upd_result = result.GetUpdateResult(0);
                resultText = this.getResultText(upd_result);
                Program.Dash.setUpdateInstalledAt(update, installBatch);
            }

            catch (Exception e)
            {
                resultText = string.Format("There was a problem installing update '{0}' : {1}", update.Title, e.Message);
                Program.Events.WriteEntry(string.Format("There was a problem installing update '{0}' : {1}", update.Title, e.Message), EventLogEntryType.Error);
            }

            return resultText;
        }

        private string getResultText( IUpdateInstallationResult upd_result )
        {
            OperationResultCode resultCode = upd_result.ResultCode;

            switch( (int) resultCode )
            {
                case 0 :
                    return "Operation Not Started";

                case 1 :
                    return "Operation In Progress";

                case 2 :
                    return "Operation Successful";

                case 3 :
                    return "Operation Completed with Errors";

                case 4 :
                    return "Operation Failed";

                case 5 :
                    return "Operation Aborted";
            }

            return "Nonstandard result code";
        }

        private string getResultMessage(List<string> results)
        {
            string message = string.Empty;

            message = "Windows Update Installation Report On <br/><br/>";

            foreach (string result in results)
            {
                message += result + "</br>";
            }

            return message;
        }
    }
}
