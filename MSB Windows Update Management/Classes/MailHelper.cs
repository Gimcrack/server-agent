using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;


namespace MSB_Windows_Update_Management
{
    public partial class MailHelper
    {
        public string mailServer = "smtp.mandrillapp.com";
        public string mailUser = "jeremy.bloomstrom@gmail.com";
        public string mailPass = "uTNf8fW-lIaod8sJ20eMtA";

        public MailHelper()
        {
        }

        public void alert(string message, string subject = null)
        {
            List<string> emails = this.GetNotificationEmails();
            
            if (subject == null)
            {
                subject = "Service Error On: " + Environment.MachineName;
            }

            foreach (string email in emails)
            {
                this.mail(message, subject, email, Environment.MachineName + "@msb.matsugov.lan");
            }

        }

        public void updateReport(string message)
        {
            List<string> emails = this.GetNotificationEmails();

            string subject = Environment.MachineName + " Update Installation Report";

            foreach (string email in emails)
            {
                this.mail(message, subject, email, Environment.MachineName + "@msb.matsugov.lan");
            }
        }

        public void mail(string message, string subject, string to, string from)
        {
            try
            {
                MailMessage m = new MailMessage();
                SmtpClient smtp = new SmtpClient(mailServer);

                m.From = new MailAddress(from);

                m.To.Add(to);

                m.Subject = subject;
                m.Body = message;
                m.IsBodyHtml = true;

                smtp.Port = 587;
                smtp.Credentials = new System.Net.NetworkCredential(mailUser, mailPass);
                smtp.EnableSsl = true;

                smtp.Send(m);
                smtp.Dispose();
            }

            catch (Exception e)
            {
                Program.Events.WriteEntry("Problem Sending Email" + Environment.NewLine + e.Source + Environment.NewLine + e.GetType().ToString() + Environment.NewLine + e.Message + Environment.NewLine + e.HResult.ToString() + Environment.NewLine + "Message" + Environment.NewLine + message + Environment.NewLine + e.StackTrace.ToString(), EventLogEntryType.Error);
            }
            
        }

        public List<string> GetNotificationEmails()
        {
            List<string> emails = new List<string>();
            SqlCommand command = new SqlCommand(Program.DB.queryGetEmailNotifications);
            
            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return emails;
            }
            else
            {
                foreach(DataRow row in rows.Rows)
                {
                    string[] tmp = row["email"].ToString().Split('\n');
                    foreach (string email in tmp)
                    {
                        if (email.Trim().Length > 0)
                        {
                            emails.Add( email.Trim() );
                        }
                    }
                }
            }

            return emails;
        }

        
    }
}
