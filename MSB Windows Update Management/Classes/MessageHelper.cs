using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Data;
using System.Data.SqlClient;

namespace MSB_Windows_Update_Management
{
    public partial class MessageHelper
    {
        string baseUrl = "https://api.tropo.com/1.0/sessions";
        string urlParams = "?action=create&token=497a625a694d775a76506e4d716e5749496d52495752624a4263586b7952614b504554584c4276516a594b4e&numberToDial={0}&msg={1}";


        public void alert(string message)
        {
            List<string> numbers = this.GetNotificationPhoneNumbers();

            foreach (string number in numbers)
            {
                this.SendTextMessage(number, message);
            }
        }

        public async void SendTextMessage(string phone, string message)
        {
            string query = string.Format(urlParams, this.formatPhone(phone), message);

            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(baseUrl);
                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

                // List data response.
                HttpResponseMessage response = await client.GetAsync(query);  // Blocking call!
                Console.WriteLine("success");
            }
        }

        public string formatPhone(string phone)
        {
            switch(phone.Length)
            {
                case 7 :
                    return "1907" + phone;

                case 10 :
                    return "1" + phone;
            }
            return phone;
        }

        public List<string> GetNotificationPhoneNumbers()
        {
            List<string> numbers = new List<string>();
            SqlCommand command = new SqlCommand(Program.DB.queryGetTextNotifications);

            DataTable rows = Program.DB.executeQuery(command);

            if (rows.Rows.Count < 1)
            {
                return numbers;
            }
            else
            {
                foreach (DataRow row in rows.Rows)
                {
                    string[] tmp = row["phone_number"].ToString().Split('\n');
                    foreach (string phone in tmp)
                    {
                        if (phone.Trim().Length > 0)
                        {
                            numbers.Add( this.formatPhone(phone) );
                        }
                    }
                }
            }

            return numbers;
        }
    }
}
