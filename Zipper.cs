using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

namespace RO_Functions
{
    class Zipper
    {
        public static List<string> ExtractZipFiles(string folderWithZipFilesForExtraction, StreamWriter sw)
        {
            DateTime method_start = DateTime.Now;
            sw.WriteLine("{0}{1}{2}", new String('#', 65), "Extraction Of Zip Files", new String('#', 85));
            List<string> extractedFolders = new List<string>();
            IEnumerable<string> all_zips = Helper.GetAllZipFiles(folderWithZipFilesForExtraction);
            sw.WriteLine("{0}Folder Path With Zip Archives: {1}", Helper.Indent(10), folderWithZipFilesForExtraction);
            sw.WriteLine("{0}Number Of Archives: {1}", Helper.Indent(10), all_zips.Count());
            sw.WriteLine("{0}Method start time: {1}", Helper.Indent(10), method_start);
            sw.WriteLine("{0}{1}", Helper.Indent(10),new String('*', 155));
            sw.WriteLine("{0}{1}\t\t{2}\t\t\t\t{3}\t\t\t\t\t\t{4}\t{5}", Helper.Indent(10), "NO", "Zip Name", "Zip Path","Start Extraction", "End Extraction");
            int counter = 1;
            foreach (string s in all_zips)
            {
                DateTime start_time = DateTime.Now;
                //string extractionFolder = Helper.GetFileNameFromPath(s).Split(new char[] { '.'})[0];
                string zipName = Helper.GetNameOfZipFile(s);
                string destinationPath = String.Format(@"{0}\{1}", folderWithZipFilesForExtraction, zipName);
                
                ZipFile.ExtractToDirectory(s, destinationPath);
                DateTime end_time = DateTime.Now;
                sw.WriteLine("{0}. {1}\t{2}\t{3}\t{4}", Helper.Indent(10) + counter++, zipName, destinationPath, start_time, end_time);
                extractedFolders.Add(destinationPath);
            }
          
            sw.WriteLine("{0}{1}", Helper.Indent(10), new String('*', 155));
            DateTime method_end= DateTime.Now;
            sw.WriteLine("{0}Method end time: {1}", Helper.Indent(10), method_end);
            sw.WriteLine("{0}{1}{2}", new String('#', 65), "End Extraction", new String('#', 94));
            sw.Close();
            return extractedFolders;
        }
    }
}
