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
            UpdateServices().Wait();
        }

        private async Task UpdateServices()
        {

            var svcs = new List<ServerService>();

            foreach(ServiceController sc in ServiceController.GetServices())
            {
                var svc = new ServiceControllerEx(sc);
                
                svcs.Add(new ServerService()
                {
                    name = svc.DisplayName.ToString(),
                    status = svc.Status.ToString(),
                    start_mode = svc.StartupType.ToString()
                });
            }

            Program.Api.UpdateServerServices(svcs);

        }
    }
}
