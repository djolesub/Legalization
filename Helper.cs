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
    class Helper
    {
        //Return List that contains all directories in specified folder
        public static IEnumerable<string> GetAllDirectories(string path = @"D:\1TestingProject")
        {
            IEnumerable<string> all_folders = null; 
            try
            {
                all_folders = Directory.EnumerateDirectories(path);
            }
            catch(Exception e)
            {
                Console.WriteLine("Error Message is: {0}", e.Message);
            }
            return all_folders ;

        }

        //Copy folder with access file and subfolder with .pdf, .doc, .docx documents to destination folder(folder for migration)
        public static void CopyDistrictFolderFilesToDestination(string destinationTopFolder, string extractedFolderPath, string sourceFolderIdPath, StreamWriter sw)
        {
            string destinationFolderFiles = Path.Combine(destinationTopFolder, extractedFolderPath.Split(Path.DirectorySeparatorChar).Last(), Helper.GetFolderIDFromFolderPath(extractedFolderPath));
            IEnumerable<string> files = null;
            sw.WriteLine("{0}Source Folder Path: {1}", Helper.Indent(10), sourceFolderIdPath);
            sw.WriteLine("{0}Destination Top Folder Path: {1}", Helper.Indent(10), destinationTopFolder);
            sw.WriteLine("{0}Destination Folder Path: {1}", Helper.Indent(10), destinationFolderFiles);

            Console.WriteLine("destination Folder: {0}", destinationFolderFiles);
            try
            {
                if (!Directory.Exists(destinationFolderFiles))
                {
                    DirectoryInfo di = Directory.CreateDirectory(destinationFolderFiles);
                    sw.WriteLine("{0}Creating Directory: {1}", Helper.Indent(10), destinationFolderFiles);
                    Console.WriteLine("Directory created:{0}", di.FullName);
                }
            }
            catch (IOException eio)
            {
                Console.WriteLine("IO Exception; {0}", eio);
                sw.WriteLine("{0}Error occured: {1}", Helper.Indent(10), eio.Message);
                sw.Close();
            }
            catch (UnauthorizedAccessException eae)
            {
                Console.WriteLine("Unautorized exception: {0}", eae.Message);
                sw.WriteLine("{0}Error occured: {1}", Helper.Indent(10), eae.Message);
                sw.Close();
            }
            finally
            {
               
            }

            try
            {
                files = Directory.EnumerateFiles(sourceFolderIdPath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while enumerating files: {0}", e.Message);
            }

            sw.WriteLine("{0}Starting Copying Files time: {1}", Helper.Indent(10), DateTime.Now);
            foreach (string filename in files)
            {
                Console.WriteLine("Filename: {0}", filename);
                try
                {
                    File.Copy(filename, Path.Combine(destinationFolderFiles, Helper.GetFileNameFromPath(filename)));
                }
                catch (Exception e)
                {
                    sw.WriteLine("{0}Error Occured: {1}", Helper.Indent(10), e.Message);
                    sw.WriteLine("{0}Method That Throws Error: {1}", Helper.Indent(10), e.TargetSite);
                    sw.Close();
                }
                finally
                {
                    
                }
            }
            sw.WriteLine("{0}Finishing Copying Files time: {1}", Helper.Indent(10), DateTime.Now);
            sw.Close();
        }

        //Copy Access file 
        public static void CopyDistrictAccessFileToDestination(string destinationTopFolder, string extractedFolderPath, string sourceAccessPath, StreamWriter sw)
        {
            string destinationFolderAccess = Path.Combine(destinationTopFolder, extractedFolderPath.Split(Path.DirectorySeparatorChar).Last());
            string destinationAccess = Path.Combine(destinationTopFolder, extractedFolderPath.Split(Path.DirectorySeparatorChar).Last(), "UMP_RR.accdb");

            sw.WriteLine("{0}Destination Top Folder: {1}", Helper.Indent(10), destinationTopFolder);
            sw.WriteLine("{0}Destination Source Access: {1}", Helper.Indent(10), sourceAccessPath);
            sw.WriteLine("{0}Destination Access File: {1}", Helper.Indent(10), destinationAccess);

            try
            {
                if (!Directory.Exists(destinationFolderAccess))
                {
                    DirectoryInfo di = Directory.CreateDirectory(destinationFolderAccess);
                    sw.WriteLine("{0}Creating Directory: {1}", Helper.Indent(10), destinationFolderAccess);
                    Console.WriteLine("Directory created:{0}", di.FullName);
                }
            }
            catch (IOException eio)
            {
                Console.WriteLine("IO Exception; {0}", eio);
                sw.WriteLine("{0}Error Occured: {1}", Helper.Indent(10), eio.Message);
                sw.WriteLine("{0}Method That Throws Error: {1}", Helper.Indent(10), eio.TargetSite);
                sw.Close();

            }
            catch (UnauthorizedAccessException eae)
            {
                Console.WriteLine("Unautorized exception: {0}", eae.Message);
                sw.WriteLine("{0}Error Occured: {1}", Helper.Indent(10), eae.Message);
                sw.WriteLine("{0}Method That Throws Error: {1}", Helper.Indent(10), eae.TargetSite);
                sw.Close();
            }
            finally
            {
                
            }

            try
            {
                File.Copy(sourceAccessPath, destinationAccess);
            }
            catch (UnauthorizedAccessException euae)
            {
                Console.WriteLine("Error occured: {0}", euae.Message);
                sw.WriteLine("{0}Error Occured: {1}", Helper.Indent(10), euae.Message);
                sw.WriteLine("{0}Method That Throws Error: {1}", Helper.Indent(10), euae.TargetSite);
                sw.Close();
            }
            catch (ArgumentException eae)
            {
                Console.WriteLine("Error occured: {0}", eae.Message);
                sw.WriteLine("{0}Error Occured: {1}", Helper.Indent(10), eae.Message);
                sw.WriteLine("{0}Method That Throws Error: {1}", Helper.Indent(10), eae.TargetSite);
                sw.Close();
            }
            catch (DirectoryNotFoundException dnf)
            {
                Console.WriteLine("Error occured: {0}", dnf.Message);
                sw.WriteLine("{0}Error Occured: {1}", Helper.Indent(10), dnf.Message);
                sw.WriteLine("{0}Method That Throws Error: {1}", Helper.Indent(10), dnf.TargetSite);
                sw.Close();
            }
            finally
            {
                
            }
            sw.Close();
        }

        //Arhiving Zip Files
        public static void PutZipFilesIntoArchiveFolder(string pathWithZipFiles, string destinationtopFolder, StreamWriter sw)
        {
            IEnumerable<string> allArrivedZipFiles = Helper.GetAllZipFiles(pathWithZipFiles);
            Console.WriteLine("Number of zip files for archiving and extraction: {0}", allArrivedZipFiles.Count());
            sw.WriteLine("{0}Starting Archiving: {1}", Helper.Indent(10), DateTime.Now);
            foreach (string s in allArrivedZipFiles)
            {

                Console.WriteLine("Zip archive for copying: {0}",s);
                string id = Helper.GetZipFileIDFromZipName(Helper.GetFileNameFromPath(s));
                Console.WriteLine("Destination Folder: {0}", Helper.GetFolderPath(id, destinationtopFolder));
                try
                {
                    File.Copy(s, Path.Combine(Helper.GetFolderPath(id, destinationtopFolder), Helper.GetFileNameFromPath(s)));
                }
                catch (Exception e)
                {
                    sw.WriteLine("{0}Error occured: {0}", Helper.Indent(10), e.Message);
                    sw.Close();
                }
                finally
                {
                    sw.Close();
                }
               
            }
            /*sw.WriteLine("{0}Ending Archiving: {1}", Helper.Indent(10), DateTime.Now);
            sw.Close();*/

        }

        //Get Folder ID Form folder name
        public static string GetFolderIDFromFolderName(string folder_name, int id_index = 0)
        {
            string folder_id = folder_name.Split(new char[] { '_'})[id_index];
            return folder_id;
        }

        //Get ZipFile ID Form folder name
        public static string GetZipFileIDFromZipName(string folder_name, int id_index = 0)
        {
            string folder_id = folder_name.Split(new char[] { '_' })[id_index];
            return folder_id;
        }

        //Get Folder ID Form folder name with path
        public static string GetFolderIDFromFolderPath(string folder_name_with_path)
        {
            string folder_name = folder_name_with_path.Split(Path.DirectorySeparatorChar).Last();
            string folder_id = GetFolderIDFromFolderName(folder_name, 0);
            return folder_id;
        }

        //Delete all files and folders from top folder. This is called after archiving all zips and write statistics to excel.
        public static void DeleteAll(string topFolderName)
        {
            IEnumerable<string> all_zips = Directory.EnumerateFiles(topFolderName);
            IEnumerable<string> all_folders = Directory.EnumerateDirectories(topFolderName);
            foreach (string filename in all_zips)
                File.Delete(filename);

            foreach (string foldername in all_folders)
                Directory.Delete(foldername);
        }

        //Get number of zip fIles in specified folder
        public static int GetNUmberOfZipFilesInFolder(string folder_name_with_path)
        {
            IEnumerable<string> all_files = Directory.GetFiles(folder_name_with_path).Where(file => file.EndsWith(".zip") || file.EndsWith(".rar"));
            return all_files.Count();
        }

        //Get all zip files from folder
        public static IEnumerable<string> GetAllZipFiles(string folder_name_with_path)
        {
            IEnumerable<string> all_files = Directory.EnumerateFiles(folder_name_with_path).Where(file => file.EndsWith(".zip") || file.EndsWith(".rar"));
            return all_files;
        }

        //Get number of fIles in specified folder
        public static int GetNUmberOfFilesInFolder(string folder_name_with_path)
        {
            string[] all_files = Directory.GetFiles(folder_name_with_path);
            return all_files.Length;
        }

        //Get File name from path
        public static string GetFileNameFromPath(string path)
        {
            return path.Split(Path.DirectorySeparatorChar).Last();
        }

        //Get id from path of specified municipality 
        public static string GetID(string folderName, char separator)
        {
            return folderName.Split(new char[] { separator }).First();
        }

        //Get namefrom path of specified municipality 
        public static string GetName(string folderName, char separator)
        {
            return folderName.Split(new char[] { separator }).Last();
        }

        //Get name of zip file 
        public static string GetNameOfZipFile(string pathToZip)
        {
            return GetFileNameFromPath(pathToZip).Split(new char[] { '.' })[0];
        }

        //Get zip arriving date from filename
        public static string GetFileArrivingDate(string filename, int index, char separator)
        {
            string date =  GetFileNameFromPath(filename).Split(new char[] { separator})[index].Split(new char[] { '-'})[0];
            if (date.Length > 7)
            {
                string year = date.Substring(0, 4);
                string month= date.Substring(4, 2);
                string day = date.Substring(6, 2);
                return String.Format("{0}.{1}.{2}", day,month,year);
               
            }
            else
            {
                return "No Date";
            }
            
        }

        //Create Zip Of Specified Folder
        public static void ZipSpecifiedFile(string fileNameWithPath)
        {
            string dest = fileNameWithPath.Split(Path.DirectorySeparatorChar).Last();
            try
            {
                string folder_id = GetFolderIDFromFolderPath(fileNameWithPath);
                string folder_path = GetFolderPath(folder_id);
                Console.WriteLine("Something: {0}", Path.Combine(folder_path, dest + ".zip"));
                ZipFile.CreateFromDirectory(fileNameWithPath, Path.Combine(folder_path, dest+".zip"));
                Console.WriteLine("GetFolderIDFromFolderName is Zipped");
            }
            catch (IOException e)
            {
                Console.WriteLine("Error: {0}", e.Message);
            }
            
        }

        //Get Absolute folder path by specified KO_ID parameter
        public static string GetFolderPath(string id, string root_folder = @"D:\1TestingProject")
        {
            IEnumerable<string> subfolders = GetAllDirectories(root_folder);
            return subfolders.Where<string>(folder => 
                                                    folder.Split(Path.DirectorySeparatorChar)
                                                    .Last()
                                                    .StartsWith(id))
                                                    .First();
        }

        public static void CreateTestFoldersAndFiles()
        {
            //Creting Files For Testing Project 
            Directory.CreateDirectory(@"D:\1TestingProject");
            for (int i = 0; i < 10; i++)
            {
                Directory.CreateDirectory(String.Format(@"D:\1TestingProject\{0}_{1}", 70500 + 100 * i, "KatOpstina"));
            }

            //Create Files in folders for testing proposes 
            IEnumerable<string> all_folders = Directory.EnumerateDirectories(@"D:\1TestingProject");
            foreach (string s in all_folders)
            {
                File.Create(Path.Combine(s, s.Split(Path.DirectorySeparatorChar).Last() + "UMP_10_20102017.zip"));
                File.Create(Path.Combine(s, s.Split(Path.DirectorySeparatorChar).Last() + "UMP_20_31052017.zip"));
                File.Create(Path.Combine(s, s.Split(Path.DirectorySeparatorChar).Last() + "UMP_30_05042017.zip"));
                File.Create(Path.Combine(s, s.Split(Path.DirectorySeparatorChar).Last() + "UMP_40_16122017.zip"));

            }
        }

        //Get number of rows in microsoft access db
        public static int GetNumberOfRowsInAccessFile(OleDbConnection conn)
        {
            int rowsReturned = 0;
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                string query = @"SELECT COUNT(ID) FROM RR";
                OleDbCommand sql = new OleDbCommand(query, conn);
                return (int)sql.ExecuteScalar();
                
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine("There was an error(InvalidOperationException): {0}.\n{1}", e.Message, e.StackTrace);
            }
            catch (OleDbException e1)
            {
                Console.WriteLine("There was an error(InvalidOperationException): {0}.\n{1}", e1.Message, e1.StackTrace);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            return rowsReturned;
        }
        //String indent
        public static string Indent(int count)
        {
            return "".PadLeft(count);
        }

        //Get number of rows in access and files in folder
        public static void GetNumberOfRowsAndFiles(string looger, string s, out int numberOfFilesInFolder, out int numberOfRowsInAccess)
        {
            StreamWriter sw1 = File.AppendText(looger);
            Console.WriteLine("#######################################################################");
            Console.WriteLine("Current folder: {0}", s);

            string folderPath = Path.Combine(s, Helper.GetID(Helper.GetFileNameFromPath(s), '_'));
            numberOfFilesInFolder = Directory.Exists(folderPath) ? Directory.EnumerateFiles(folderPath).Where<string>(
                                                                                                    doc => doc.EndsWith(".pdf") ||
                                                                                                    doc.EndsWith(".doc") ||
                                                                                                    doc.EndsWith(".docx")
                                                                                                ).Count() : 0;

            sw1.WriteLine("{0}{1}", Helper.Indent(10), new String('*', 155));

            string accessPath = Path.Combine(s, @"UMP_RR.accdb");
            string connectionStringForAccess = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source ={0};", accessPath);
            OleDbConnection conn1 = new OleDbConnection(connectionStringForAccess);
            numberOfRowsInAccess = File.Exists(accessPath) ? Helper.GetNumberOfRowsInAccessFile(conn1) : 0;
            Console.WriteLine("Access Path: {0}", accessPath);

            sw1.WriteLine("{0}Folder Path: {1}", Helper.Indent(10), folderPath);
            sw1.WriteLine("{0}Files In Folder: {1}", Helper.Indent(10), numberOfFilesInFolder);
            sw1.WriteLine("{0}Access Path: {1}", Helper.Indent(10), accessPath);
            sw1.WriteLine("{0}Rows In Access: {1}", Helper.Indent(10), numberOfRowsInAccess);
            sw1.WriteLine("{0}Access Connection: {1}", Helper.Indent(10), connectionStringForAccess);
            //sw1.WriteLine("{0}{1}", Helper.Indent(10), new String('*', 155));
            Console.WriteLine("Number of files in folder: {0},Access:{1},Parh:{2}", numberOfFilesInFolder, numberOfRowsInAccess, folderPath);
            sw1.Close();
        }

        //WriteInformationsToExcel
        public static void WriteInformationsToExcel(string looger, IEnumerable<string> extracted_folders, string pathOfExcel, string sheetName)
        {
            foreach (string s in extracted_folders)
            {
                
                int numberOfFilesInFolder = 0;
                int numberOfRowsInAccess = 0;
                Helper.GetNumberOfRowsAndFiles(looger, s, out numberOfFilesInFolder, out numberOfRowsInAccess);

                /******************************************Working with excel***********************************************/
                StreamWriter sw4 = File.AppendText(looger);
                int maticni = Convert.ToInt32(Helper.GetFolderIDFromFolderPath(s));
                OleDbConnection excelConn = new OleDbConnection(String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;", pathOfExcel));

                MyExcel.UpdateOrInsert(excelConn, maticni, numberOfFilesInFolder, numberOfRowsInAccess, sheetName, sw4);
                Console.WriteLine("Finishing import for KO: {0}", s);
            }
        }

        // Get Error types
        public static int GetErrorType(int numberOfFiles, int numberOfRows)
        {
            
            if (numberOfFiles == 0)
                return 2;
            else if (numberOfRows == 0)
                return 3;
            else if ((numberOfFiles - numberOfRows) != 0)
                return 1;
            return 0;
        }

        //Put to migration folder 
        public static void MoveDataToFoldersForMigration(string looger, IEnumerable<string> extracted_folders, string consistentFolder, string noConsistentFolder)
        {
            foreach (string s in extracted_folders)
            {
                StreamWriter sw6 = File.AppendText(looger);
                string sourceAccessPath = Path.Combine(s, "UMP_RR.accdb");
                string sourceFolderIdPath = Path.Combine(s, Helper.GetFolderIDFromFolderPath(s));

                sw6.WriteLine("{0}{1}", Helper.Indent(10), new String('*', 155));
                sw6.WriteLine("{0}Folder Path: {1}", Helper.Indent(10), s);
                sw6.WriteLine("{0}Access For Copying Path: {1}", Helper.Indent(10), sourceAccessPath);
                sw6.WriteLine("{0}Folder With Files For Copying: {1}", Helper.Indent(10), sourceFolderIdPath);
                sw6.Close();

                if (Directory.Exists(sourceFolderIdPath) && File.Exists(sourceAccessPath))
                {
                    StreamWriter sw8 = File.AppendText(looger);
                    Helper.CopyDistrictFolderFilesToDestination(consistentFolder, s, sourceFolderIdPath, sw8);

                    StreamWriter sw9 = File.AppendText(looger);
                    Helper.CopyDistrictAccessFileToDestination(consistentFolder, s, sourceAccessPath, sw9);

                    StreamWriter sw10 = File.AppendText(looger);
                    Helper.CopyDistrictAccessFileToDestination(noConsistentFolder, s, sourceAccessPath, sw10);
                }

                else if (File.Exists(sourceAccessPath))
                {
                    StreamWriter sw11 = File.AppendText(looger);
                    Helper.CopyDistrictAccessFileToDestination(noConsistentFolder, s, sourceAccessPath, sw11);
                }
                StreamWriter sw13 = File.AppendText(looger);
                //sw13.WriteLine("{0}{1}", Helper.Indent(10), new String('*', 155));
                sw13.Close();
            }
        }
    }
}
