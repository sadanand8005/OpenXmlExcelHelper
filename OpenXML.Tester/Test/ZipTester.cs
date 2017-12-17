using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML.Tester.Test
{
    public class ZipTester
    {

        public void Test()
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test");
            string zipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ABC.xlsx");

            ZipFile.CreateFromDirectory(folder, zipPath, CompressionLevel.Fastest, false);
        }

        public void Test2()
        {
            //ZipOutputStream zip = new ZipOutputStream(File.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ABC.xlsx")));

            //zip.SetLevel(9);

            //string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Test");

            //ZipFolder(folder, folder, zip);
            //zip.Finish();
            //zip.Close();

        }

        //private void ZipFolder(string RootFolder, string CurrentFolder, ZipOutputStream zStream)
        //{
        //    string[] SubFolders = Directory.GetDirectories(CurrentFolder);

        //    foreach (string Folder in SubFolders)
        //        ZipFolder(RootFolder, Folder, zStream);

        //    string relativePath = CurrentFolder.Substring(RootFolder.Length) + "\\";

        //    if (relativePath.Length > 1)
        //    {
        //        ZipEntry dirEntry;

        //        dirEntry = new ZipEntry(relativePath);
        //        dirEntry.DateTime = DateTime.Now;
        //    }

        //    foreach (string file in Directory.GetFiles(CurrentFolder))
        //    {
        //        AddFileToZip(zStream, relativePath, file);
        //    }
        //}

        //private void AddFileToZip(ZipOutputStream zStream, string relativePath, string file)
        //{
        //    byte[] buffer = new byte[4096];
        //    string fileRelativePath = (relativePath.Length > 1 ? relativePath : string.Empty) + Path.GetFileName(file);
        //    ZipEntry entry = new ZipEntry(fileRelativePath);

        //    entry.DateTime = DateTime.Now;
        //    zStream.PutNextEntry(entry);

        //    using (FileStream fs = File.OpenRead(file))
        //    {
        //        int sourceBytes;

        //        do
        //        {
        //            sourceBytes = fs.Read(buffer, 0, buffer.Length);
        //            zStream.Write(buffer, 0, sourceBytes);
        //        } while (sourceBytes > 0);
        //    }
        //}

    }
}
