using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML.Tester.Test
{
    public class RandomData
    {
        public int Id { get; set; }
        public string Field1 { get; set; }
        public string Field2 { get; set; }
        public string Field3 { get; set; }
        public string Field4 { get; set; }
        public string Field5 { get; set; }
        public string Field6 { get; set; }
        public string Field7 { get; set; }
        public string Field8 { get; set; }
        public string Field9 { get; set; }
        public string Field10 { get; set; }
        public string Field11 { get; set; }
        public string Field12 { get; set; }
        public string Field13 { get; set; }
        public string Field14 { get; set; }
        public string Field15 { get; set; }
        public string Field16 { get; set; }
        public string Field17 { get; set; }
        public string Field18 { get; set; }
        public string Field19 { get; set; }
    }

    public class DataHelper
    {
        public void CreateData()
        {
            var random = new Random();
            using (StreamWriter writer = new StreamWriter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sample.json")))
            {
                for (int i = 0; i < 50000; i++)
                {
                    var obj = new RandomData();
                    obj.Id = random.Next();
                    obj.Field1 = GetRandomString();
                    obj.Field2 = GetRandomString();
                    obj.Field3 = GetRandomString();
                    obj.Field4 = GetRandomString();
                    obj.Field5 = GetRandomString();
                    obj.Field6 = GetRandomString();
                    obj.Field7 = GetRandomString();
                    obj.Field8 = GetRandomString();
                    obj.Field9 = GetRandomString();
                    obj.Field10 = GetRandomString();
                    obj.Field11 = GetRandomString();
                    obj.Field12 = GetRandomString();
                    obj.Field13 = GetRandomString();
                    obj.Field14 = GetRandomString();
                    obj.Field15 = GetRandomString();
                    obj.Field16 = GetRandomString();
                    obj.Field17 = GetRandomString();
                    obj.Field18 = GetRandomString();
                    obj.Field19 = GetRandomString();

                    writer.Write(JsonConvert.SerializeObject(obj));
                }
            }
        }

        public List<RandomData> GetData()
        {
            var data = new List<RandomData>();
            JsonSerializer ser = new JsonSerializer();

            using (StreamReader reader = new StreamReader(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sample.json")))
            {
                using (var jsonReader = new JsonTextReader(reader))
                {
                    while (jsonReader.Read())
                    {
                        if (jsonReader.TokenType == JsonToken.StartObject)
                        {
                            var obj = ser.Deserialize<RandomData>(jsonReader);

                            data.Add(obj);
                        }
                    }
                    
                }
            }

            return data;
        }

        private string GetRandomString()
        {
            string path = Path.GetRandomFileName();
            path = path.Replace(".", ""); // Remove period.
            return path;
        }

    }
}
