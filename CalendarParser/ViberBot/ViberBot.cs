using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;

namespace ViberBot
{
    public class ViberBot
    {
        public void Run() {
            var request = WebRequest.Create(new Uri("https://chatapi.viber.com/pa/get_account_info"));
            request.Headers.Add("X-Viber-Auth-Token", "46c1e71c3fe7d503-2f0a86b2d13379cd-cbd469e6a3d06886");
            request.Method = "POST";
            
            var response = request.GetResponse();
            var stream = response.GetResponseStream();
            using (var reader = new StreamReader(stream))
            {
                var data = reader.ReadToEnd();
            }

        }
    }
}
