using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Management;

namespace MSB_Windows_Update_Management
{
    public class ServiceControllerEx
    {
        public ServiceController ServiceController;

        public ServiceControllerEx(ServiceController sc)
        {
            this.ServiceController = sc;
        }

        public string ServiceName
        {
            get
            {
                return this.ServiceController.ServiceName;
            }
        }

        public string DisplayName
        {
            get
            {
                return this.ServiceController.DisplayName;
            }
        }

        public ServiceControllerStatus Status
        {
            get
            {
                return this.ServiceController.Status;
            }
        }
        

        public string Description
        {
            get
            {
                string path = "Win32_Service.Name='" + this.ServiceController.ServiceName + "'";
                ManagementPath p = new ManagementPath(path);
                ManagementObject m = new ManagementObject(p);
                if (m["Description"] != null)
                {
                    return m["Description"].ToString();
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        public string StartupType
        {
            get
            {
                string path = "Win32_Service.Name='" + this.ServiceController.ServiceName + "'";
                ManagementPath p = new ManagementPath(path);
                ManagementObject m = new ManagementObject(p);
                if (m["Description"] != null)
                {
                    return m["StartMode"].ToString();
                }
                else
                {
                    return string.Empty;
                }  
            }
        }
    }
}
