using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML.Tester.Test
{
    class TestDataHelper
    {
        public List<TestData> GetData()
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://jsonplaceholder.typicode.com/photos");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                
                HttpResponseMessage response = client.GetAsync("").GetAwaiter().GetResult();
                if (response.IsSuccessStatusCode)
                {
                    List<TestData> data = response.Content.ReadAsAsync<List<TestData>>().Result;
                    return data;
                }
                else
                {
                    Console.WriteLine("Internal server Error");
                }
            }

            return null;
        }
    }

    public class TestData
    {
        public int albumId { get; set; }
        public int id { get; set; }
        public string title { get; set; }
        public string url { get; set; }
        public string thumbnailUrl { get; set; }
    }
}
