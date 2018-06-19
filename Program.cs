using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Data;
using System.Data.OleDb;

namespace RO_Functions
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //NewMethod();
            //ZipFile.ExtractToDirectory(@"\\PROCENA-08\Srki\9 Resenja\2 Kontrola\70556_UMPRR_20180529-144137.zip", @"\\PROCENA-08\Srki\9 Resenja\2 Kontrola");
            string looger = String.Format("{0}_{1}{2}", @"\\PROCENA-08\Srki\9 Resenja\IspravkagreskiNeKonzistentno", DateTime.Now.ToString("dd.MM.yyy"), ".txt");
            Console.WriteLine(looger);
            StreamWriter sw = File.AppendText(looger);

            List<string> extracted_folders = Zipper.ExtractZipFiles(@"\\PROCENA-08\Srki\9 Resenja\2 Kontrola", sw);
            StreamWriter sw2 = File.AppendText(looger);

            /************************Loging Informations To File*****************************************/
            sw2.WriteLine();
            sw2.WriteLine("{0}{1}{2}", new String('#', 65), "Writing Information To Excel", new String('#', 80));
            sw2.Close();
            /*************************End Loging Informations to File*************************************/

/******************************************Working with excel non consistent***********************************************/
            string pathOfExcel = @"C:\Users\djordje.subotic\Desktop\NeBrisi\Evidencija.xlsx";
            string sheetName = "11.jun";/*Sheet name Djole - zadnji*/
            Helper.WriteInformationsToExcel(looger, extracted_folders, pathOfExcel, sheetName);
            
            /*************************Loging Informations To File*****************************************/
            StreamWriter sw3 = File.AppendText(looger);
            sw3.WriteLine("{0}{1}{2}", new String('#', 65), "End Writing", new String('#', 96));
            sw3.Close();

            StreamWriter sw5 = File.AppendText(looger);
            sw5.WriteLine();
            sw5.WriteLine("{0}{1}{2}", new String('#', 65), "Copying Extracted Folders To Migration Folder", new String('#', 80));
            sw5.Close();
            /*************************End Loging Informations to File*************************************/
            Helper.MoveDataToFoldersForMigration(looger, extracted_folders, @"\\PROCENA-08\Srki\9 Resenja\4 Konzistento", @"\\PROCENA-08\Srki\9 Resenja\3 Djole");

            /*************************Loging Informations To File*****************************************/
            StreamWriter sw7 = File.AppendText(looger);
            sw7.WriteLine("{0}{1}{2}", new String('#', 65), "End Copying", new String('#', 96));
            sw7.Close();

            StreamWriter sw12 = File.AppendText(looger);
            sw12.WriteLine();
            sw12.WriteLine("{0}{1}{2}", new String('#', 65), "Copying Zip Files To Archive Folder", new String('#', 80));
            sw12.Close();
            StreamWriter sw14 = File.AppendText(looger);
            /*************************End Loging Informations to File*************************************/
            //Put all zip archives from 2Kontrola folder to 2VecUnzipovano folder.
            Helper.PutZipFilesIntoArchiveFolder(@"\\PROCENA-08\Srki\9 Resenja\2 Kontrola", @"\\srvfile1\Resenja o rusenju\2VecUnZipovano", sw14);

            StreamWriter sw15 = File.AppendText(looger);
            sw15.WriteLine("{0}{1}{2}", new String('#', 65), "End Copying", new String('#', 96));
            sw15.Close();

            /******************************************Working with excel consistent***********************************************/
            Console.WriteLine("Waiting to remove errors....");
            Console.Write("If you are done with removing errors type [Y]:\t");
            //StreamWriter sw20 = File.AppendText(String.Format("{0}_{1}{2}", @"\\PROCENA-08\Srki\9 Resenja\IspravkagreskiKonzistentno", DateTime.Now.ToString("dd.MM.yyy"), ".txt"));
            IEnumerable<string> extracted_folders_consistent = Directory.EnumerateDirectories(@"\\PROCENA-08\Srki\9 Resenja\4 Konzistento");

            string d = Console.ReadLine();
            if (d.ToUpper() == "Y")
            {
                string loogerConsistent = String.Format("{0}_{1}{2}", @"\\PROCENA-08\Srki\9 Resenja\IspravkagreskiKonzistentno", DateTime.Now.ToString("dd.MM.yyy"), ".txt");
                string pathOfExcelConsistent = @"\\PROCENA-08\Srki\9 Resenja\EvidencijaKonzistentnog.xlsx";
                string sheetNameConsistent = "11.jun"; // sheet za konzistentno
                Helper.WriteInformationsToExcel(loogerConsistent, extracted_folders_consistent, pathOfExcelConsistent, sheetNameConsistent);
            }
        }
        
    }
}
