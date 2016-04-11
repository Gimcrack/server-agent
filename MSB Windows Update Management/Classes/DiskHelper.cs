using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MSB_Windows_Update_Management
{
    public partial class DiskHelper
    {
        public void GetInfo()
        {
            foreach (DriveInfo drive in DriveInfo.GetDrives() )
            {
                if ( drive.IsReady && ( drive.DriveType == DriveType.Fixed || drive.DriveType == DriveType.Removable ) )
                {
                    Program.Dash.storeServerDiskInDashboard(drive);
                }
            }
        }
    }
}
