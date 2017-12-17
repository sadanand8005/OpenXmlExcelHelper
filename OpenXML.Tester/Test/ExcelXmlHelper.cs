using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML.Tester.Test
{
    public class ExcelXmlHelper
    {
        public async Task<bool> Test()
        {
            using (var filestream = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test", "xl", "worksheets", "sheet1.xml"), FileMode.OpenOrCreate))
            {
                using(var writer = new StreamWriter(filestream))
                {
                    var sb = new StringBuilder();

                    AddHeaderToXmlFile(sb);
                    OpenWorkSheet(sb);
                    OpenSheetData(sb);

                    var writerTask = writer.WriteAsync(sb.ToString());
                    sb.Clear();

                    var dataHelper = new DataHelper();
                    var data = dataHelper.GetData();

                    int k = 0;

                    foreach (RandomData item in data)
                    {
                        OpenRow(sb);
                        AddCell(sb, item.Id);
                        AddCell(sb, item.Field1);
                        AddCell(sb, item.Field2);
                        AddCell(sb, item.Field3);
                        AddCell(sb, item.Field4);
                        AddCell(sb, item.Field5);
                        AddCell(sb, item.Field6);
                        AddCell(sb, item.Field7);
                        AddCell(sb, item.Field8);
                        AddCell(sb, item.Field9);
                        AddCell(sb, item.Field10);
                        AddCell(sb, item.Field11);
                        AddCell(sb, item.Field12);
                        AddCell(sb, item.Field13);
                        AddCell(sb, item.Field14);
                        AddCell(sb, item.Field15);
                        AddCell(sb, item.Field16);
                        AddCell(sb, item.Field17);
                        AddCell(sb, item.Field18);
                        AddCell(sb, item.Field19);
                        CloseRow(sb);

                        if(++k == 100)
                        {
                            k = 0;
                            await writerTask;
                            writerTask = writer.WriteAsync(sb.ToString());
                            sb.Clear();
                        }
                    }

                    CloseSheetData(sb);
                    CloseWorkSheet(sb);

                    await writerTask;
                    writerTask = writer.WriteAsync(sb.ToString());

                    await writerTask;
                }
            }

            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test");
            string zipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ABC.xlsx");

            File.Delete(zipPath);

            ZipFile.CreateFromDirectory(folder, zipPath, CompressionLevel.Fastest, false);

            return true;
        }

        private void AddHeaderToXmlFile(StringBuilder sb)
        {
            sb.Append("<?xml version='1.0' encoding='utf-8'?>");
        }

        private void OpenWorkSheet(StringBuilder sb)
        {
            sb.Append("<x:worksheet xmlns:x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>");
        }

        private void CloseWorkSheet(StringBuilder sb)
        {
            sb.Append("</x:worksheet>");
        }

        private void OpenSheetData(StringBuilder sb)
        {
            sb.Append("<x:sheetData>");
        }

        private void CloseSheetData(StringBuilder sb)
        {
            sb.Append("</x:sheetData>");
        }

        private void OpenRow(StringBuilder sb)
        {
            sb.Append("<x:row>");
        }

        private void CloseRow(StringBuilder sb)
        {
            sb.Append("</x:row>");
        }

        private void AddCell(StringBuilder sb, string value)
        {
            sb.Append("<x:c t='str'><x:v>");
            sb.Append(value);
            sb.Append("</x:v></x:c>");
        }

        private void AddCell(StringBuilder sb, int value)
        {
            sb.Append("<x:c t='n'><x:v>");
            sb.Append(value);
            sb.Append("</x:v></x:c>");
        }

        private void AddCell(StringBuilder sb, decimal value)
        {
            sb.Append("<x:c t='n'><x:v>");
            sb.Append(value);
            sb.Append("</x:v></x:c>");
        }
    }
}
