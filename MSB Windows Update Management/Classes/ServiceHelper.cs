using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Management;
using System.Net.Mail;

namespace MSB_Windows_Update_Management
{
    public partial class ServiceHelper
    {   
        public void CheckServices()
        {
            List<ServiceControllerEx> scProblem = this.GetProblemServices();

            if (scProblem.Count > 0)
            {
                string errorMessage = this.GetErrorMessage(scProblem);
                Program.Alert.ServerAlert(errorMessage);
            }
        }
        
        public string GetErrorMessage(List<ServiceControllerEx> scProblem)
        {
            string err = "Error, the following automatic services are offline on server: " + Environment.MachineName + Environment.NewLine;

            foreach(ServiceControllerEx sc in scProblem)
            {
                err += "   " + sc.DisplayName + Environment.NewLine;
            }

            err += Environment.NewLine + "--Reported at " + DateTime.Now;

            return err;
        }
        
        public List<ServiceControllerEx> GetAutomaticServices()
        {
            List<ServiceControllerEx> scAutomatic = new List<ServiceControllerEx>();
            ServiceController[] scServices = ServiceController.GetServices();
            
            foreach( ServiceController sc in scServices )
            {
                ServiceControllerEx tmpSc = new ServiceControllerEx(sc);

                if ( tmpSc.StartupType.IndexOf("Auto") != -1 )
                {
                    scAutomatic.Add(tmpSc);
                }
            }
            return scAutomatic;
        }

        public List<ServiceControllerEx> GetProblemServices()
        {
            List<ServiceControllerEx> scProblem = new List<ServiceControllerEx>();
            List<ServiceControllerEx> scAutomatic = this.GetAutomaticServices();

            foreach( ServiceControllerEx sc in scAutomatic )
            {
                if ( this.IgnoreService(sc) )
                {
                    continue;
                }

                // if an automatic service is not running, that is a problem
                if ( sc.Status != ServiceControllerStatus.Running)
                {
                    scProblem.Add(sc);
                }
            }

            return scProblem;
        }

        public bool IgnoreService(ServiceControllerEx sc)
        {
            List<string> ignoredServices = Program.Dash.GetExemptServices();

            foreach (string svc in ignoredServices)
            {
                if (sc.DisplayName.IndexOf(svc) != -1 || svc == "All" || svc == "*")
                    return true;
            }

            return false;
        }
        
    }
}
