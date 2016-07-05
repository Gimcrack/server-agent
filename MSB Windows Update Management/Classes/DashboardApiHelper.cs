using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Formatting;

namespace MSB_Windows_Update_Management
{
    public partial class DashboardApiHelper
    {
        private string url = "http://itdashboard/api/v1/";
        private string url_params = "?api_token=eoaDZD8mM7uWzuO0k5dxIInkERpplnVXDfHI0u1GeqfPsH8VQe7kWCuzJOtl";


        public async Task UpdateServerServices(List<ServerService> svcs)
        {
            using (HttpClient http = new HttpClient())
            {
                http.BaseAddress = new Uri(url);

                http.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json")
                );

                string endpoint = "Server/" + Program.Dash.getServerId().ToString() + "/Services" + url_params;

                HttpResponseMessage response = await http.PostAsJsonAsync(endpoint, svcs);
            }   
        }

        

    }
}
