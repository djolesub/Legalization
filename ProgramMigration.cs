using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Npgsql;

namespace Resenja_Folderi_Opstine
{
    class Program
    {
        static void Main(string[] args)
        {
           
            //get current directory and display it.
            string currentDirectory = Directory.GetCurrentDirectory();
            Console.WriteLine("Current Directory is {0}", currentDirectory);

            //Create root folder that we will put all other folders
            DirectoryInfo d = CreateFolderForAllPoliticalRegions(@"\\srvfile1\Resenja o rusenju\5RR");

            //Read ID's and Names from specified csv file with political districts
            ReadIDAndPoliticalDistrictNameFromCsv(@"D:\Resenja_Rusenje\Opstina.csv");


            IEnumerable a = GetAllSubFolders();
            IEnumerable b = GetAllPoliticalDistrictFolders();
            MatchUnzipNameWithPoliticalDistrictFolderName(a, b);

  
    }

        //Create root folder
        public static DirectoryInfo CreateFolderForAllPoliticalRegions(string path)
        {
            DirectoryInfo regions_folder = null;
            try
            {
                regions_folder = Directory.CreateDirectory(path);
                Console.WriteLine("Directory with name: {0} and full path: {1} is succesfuly created.", regions_folder.Name, regions_folder.FullName);
            }
            catch (ArgumentException)
            {
                Console.WriteLine("Sory, specified path is not valid.");
                Console.WriteLine("Please specified valid path for your folder");
            }

            return regions_folder;
        }

        //Read data from CSV file 
        public static void ReadIDAndPoliticalDistrictNameFromCsv(string pathToCsv)
        {
            string[] data = null;
            try
            {
                data = File.ReadAllLines(pathToCsv, Encoding.GetEncoding(1252));
                foreach (string s in data)
                {
                    Console.WriteLine(s);
                    string[] string_array = s.Split(new char[] { ',' });
                    string maticniBrojPolitickeOpstine = string_array[0];
                    string nazivPolitickeOpstine = string_array[1];
                    Console.WriteLine("ID Opstine:{0} Naziv:{1}", maticniBrojPolitickeOpstine, nazivPolitickeOpstine);
                    string full_name = maticniBrojPolitickeOpstine + "_" + nazivPolitickeOpstine;
                    CreateFolderForPoliticalDistrict(@"\\srvfile1\Resenja o rusenju\5RR", full_name);
                }

            }
            catch (DirectoryNotFoundException d)
            {
                Console.WriteLine("Sory there was an error:{0}", d.Message);
            }
            catch (FileNotFoundException f)
            {
                Console.WriteLine("Sory there was an error:{0}", f.Message);

            }
        }
        //Create Folder for each political district
        public static void CreateFolderForPoliticalDistrict(string rootFolder, string folderName)
        {
            string string_path = Path.Combine(rootFolder, folderName);
            Directory.CreateDirectory(string_path);

        }

        //List all zip files in Podaci_sa_maila directory from  srvfile server
        public static List<string> GetAllZipFiles(string path = @"\\srvfile1\Resenja o rusenju\4ZaKopiranjeuRR")
        {
            List<string> all_zip_files = new List<string>();
            try
            {

                foreach (string s in Directory.EnumerateFiles(path))
                {
                    //all_zip_files.Add(Path.GetFileName(s));

                    all_zip_files.Add(s);
                }

            }
            catch (ArgumentException e)
            {
                Console.WriteLine("There was an error with method arguments.\n{0}", e.Message);
            }
            catch (DirectoryNotFoundException d)
            {
                Console.WriteLine("Therer was an error with specified directory.\n{0}", d.Message);
            }
            return all_zip_files;
        }

        //List all folder names in RR folder from srvfile server 
        public static IEnumerable GetAllPoliticalDistrictFolders(string path = @"\\srvfile1\Resenja o rusenju\5RR")
        {
            List<string> all_district_folders = new List<string>();
            try
            {

                foreach (string h in Directory.EnumerateDirectories(path))
                {
                    //all_district_folders.Add(h.Split(Path.DirectorySeparatorChar).Last());
                    all_district_folders.Add(h);
                }
            }
            catch (ArgumentException a)
            {
                Console.WriteLine("There was an error with directory path.\n{0}", a.Message);
            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("There was an error.\n{0}", de.Message);
            }

            return all_district_folders;
        }

        //Copy Files To concrete Folder 
        public static void MatchUnzipNameWithPoliticalDistrictFolderName(IEnumerable listOfUnZipFolders, IEnumerable listOfFolders)
        {
            foreach (string politicalDistrictFolder in listOfFolders)
            {
                //Console.WriteLine("Political Name is: {0}", politicalDistrictFolder);
                foreach (string unzipFolderName in listOfUnZipFolders)
                {
                    //Console.WriteLine("#Unzip Folder Name: {0}", unzipFolderName);
                    //Get FolderName
                    string unzip_folder_last_part = unzipFolderName.Split(Path.DirectorySeparatorChar).Last();
                    string unzip_folder_id = unzip_folder_last_part.Split('_')[0];

                    //Console.WriteLine("#Unzip ID:{0}", unzip_folder_id);
                    string political_folder_name = politicalDistrictFolder.Split(Path.DirectorySeparatorChar).Last();
                    string political_folder_name_id = political_folder_name.Split(new char[] { '_' })[0];
                    //Console.WriteLine("Unzip: {0}_____Political: {1}", unzip_folder_id, political_folder_name_id);
                    if (political_folder_name_id == unzip_folder_id)
                    {
                        //Console.WriteLine("Its a match for {0}", unzipFolderName);
                        //Console.WriteLine("UnzipFoldername:{0}\nPoliticalDistrictFolder:{1}\nDistrictID:{2}", unzipFolderName, politicalDistrictFolder, political_folder_name_id);
                        CopyFilesToPoliticalDistrictFolders(unzipFolderName, politicalDistrictFolder, political_folder_name_id);
                        
                    }

                }
            }
        }
        //Check if directory exist in specifed directory with path. Does unziped folder contains folder with pdf's or word docs
        public static bool DoesDirectoryWithFilesExistInPath(string fromPath, string political_folder_name_id)
        {
            bool exist = true;
            try
            {
                exist = Directory.Exists(String.Format("{0}\\{1}", fromPath, political_folder_name_id));
            }
            catch(Exception e)
            {
                Console.WriteLine("There was an error.\n{0}",e.Message);
            }
            return exist;
        }
        //Check if MS Access file exist in specified directory with path. Does unzip folder contains MS Access file 
        public static bool DoesonlyOneAccessFileExistInPath(string toPath)
        {
            string[] access_files_array;
            bool exist = false;
            try
            {
                access_files_array = Directory.GetFiles(toPath, "*.accdb");
                if (access_files_array.Length == 1)
                {
                    exist = true;
                }

            }
            catch(DirectoryNotFoundException de)
            {
                Console.WriteLine("There was an error.\n{0}", de.Message);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("There was an error.\n{0}", ane.Message);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("There was an error.\n{0}", ae.Message);
            }

            return exist;
            
        }

        //Get Access File 
        public static string[] GetAccessFileFromPath(string fromPath)
        {/*Menjano jer puca.*/
            string[] access_files_array = null;
            

            try
            {
                access_files_array = Directory.GetFiles(fromPath, "*.accdb");
               

            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("There was an error.\n{0}", de.Message);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("There was an error.\n{0}", ane.Message);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("There was an error.\n{0}", ae.Message);
            }

            return access_files_array;
        }

        //Delete Existing Access File For mprevious Version 
        public static void DeleteExistingAccessFileFromPreviousVersion(string access_file_from_previous_version, string deleted_file_log_name= @"C:\Users\djordje.subotic\Desktop\Obrisani.txt")
        {
            Console.WriteLine("Access file {0} allready exist.", access_file_from_previous_version);
            StreamWriter deleted_file_writer = null;
           
            try
            {
                File.Delete(access_file_from_previous_version);
                try
                {
                    deleted_file_writer = File.AppendText(deleted_file_log_name);
                    Console.WriteLine("I am deleting access file now.Its Deleted");
                    deleted_file_writer.WriteLine(String.Format("{0}{1}{2}",new String('#',20),DateTime.Now.ToString(), new String('#', 20)));
                    deleted_file_writer.WriteLine("Delete file is:\t{0}", access_file_from_previous_version);
                    deleted_file_writer.Close();
                }
                catch (DirectoryNotFoundException de)
                {
                    Console.WriteLine("An error occured.\n{0}", de.Message);
                }
                catch (ArgumentNullException ane)
                {
                    Console.WriteLine("An error occured.\n{0}", ane.Message);
                }
                catch (ArgumentException ae)
                {
                    Console.WriteLine("An error occured.\n{0}", ae.Message);
                }
                finally
                {
                    if (deleted_file_writer != null)
                    {
                        deleted_file_writer.Close();
                    }
                }


            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("An error occured.\n{0}", de.Message);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("An error occured.\n{0}", ane.Message);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("An error occured.\n{0}", ae.Message);
            }
        }

        public static void CreateDirectoryAndCopyDocuments(string directory_for_files, string[] files_in_directory)
        {
            Console.WriteLine("Deleting all .pdf or .docx files form directory.");
            
            Directory.CreateDirectory(directory_for_files);
            foreach (string file_name in files_in_directory)
            {
                Console.WriteLine("Files Dir Exists are: {0}", file_name);
                string[] array_of_file_name_parts = file_name.Split(new char[] { '\\' });
                string n1 = array_of_file_name_parts[array_of_file_name_parts.Length - 1];
                Console.WriteLine("Copy new file to directory");
                try
                {
                    File.Copy(file_name, directory_for_files + "\\" + n1);
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occured.\n{0}", e.Message);
                }
            }
        }

        //Get Access File Name without path. For example: me.accdb;
        public static string GetAccessFileFromPathWithoutPath(string access_file_name_with_path)
        {
            string[] array_of_separated_parts = access_file_name_with_path.Split(new char[] { '\\' });
            string access_file_name = array_of_separated_parts[array_of_separated_parts.Length - 1];
            return access_file_name;
        }

        //Delete All Files After Coping And Migrating to database
        public static void DeleteCopiedFolder(string path)
        {
           string[] d =  Directory.GetDirectories(path);
            try
            {
                foreach (string s in d)
                {
                    Directory.Delete(s, true);
                }
                //sw.WriteLine($"Directory {path} is deleted.");
            }
            catch (ArgumentNullException ae)
            {
                Console.WriteLine("There was an error.\n{0}", ae.Message);
                //sw.WriteLine($"There was an error while deleting directory {path}");
            }
            catch (ArgumentException ane)
            {
                Console.WriteLine("There was an error.\n{0}", ane.Message);
                //sw.WriteLine($"There was an error while deleting directory {path}");
            }
            catch (IOException ioe)
            {
                Console.WriteLine("There was an error.\n{0}", ioe.Message);
                //sw.WriteLine($"There was an error while deleting directory {path}");
            }


        }

        //Copy MS Access File to apropriate location(political district folder) 
        public static string CopyAccessFileToPoliticalDistrictFolder(string access_file_with_path_name,string toPath, string access_file_name)
        {
            string copied_file_location = String.Format("{0}\\{1}", toPath, access_file_name);
            try
            {
                File.Copy(access_file_with_path_name, copied_file_location);
            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("An error occured.\n{0}", de.Message);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("An error occured.\n{0}", ane.Message);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("An error occured.\n{0}", ae.Message);
            }
            catch (FileNotFoundException fe)
            {
                Console.WriteLine("An error occured.\n{0}", fe.Message);
            }
            finally
            {

            }
            return copied_file_location;
        }


        //Check Number of Documents in folder with word and pdf 
        public static void GetNumberOfFilesInDirectory(string fromPath, ref string[] folder_with_pdf_or_word, ref string[] files_in_Directory, ref int? number_of_files_in_Directory)
        {
            try
            {
                folder_with_pdf_or_word = Directory.GetDirectories(fromPath);
                
                if (folder_with_pdf_or_word.Length > 0)
                {
                    try
                    {
                        files_in_Directory = Directory.GetFiles(folder_with_pdf_or_word[0]);
                        number_of_files_in_Directory = files_in_Directory.Length;
                    }
                    catch (DirectoryNotFoundException de)
                    {
                        Console.WriteLine("An error occured.\n{0}", de.Message);
                    }
                    catch (ArgumentNullException ane)
                    {
                        Console.WriteLine("An error occured.\n{0}", ane.Message);
                    }
                    catch (ArgumentException ae)
                    {
                        Console.WriteLine("An error occured.\n{0}", ae.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("An error occured.\n{0}", de.Message);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("An error occured.\n{0}", ane.Message);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("An error occured.There is no files in folder.\n{0}",ae.Message);
            }
           
        }

        //Copy folders to political distrivct folders 
        public static void CopyFilesToPoliticalDistrictFolders(string fromPath, string toPath, string political_folder_name_id)
        {
            Console.WriteLine("TO PATH is: {0}", toPath);
            StreamWriter sw = null;
            try
            {
                sw = File.AppendText(String.Format("{0}_{1}.{2}",@"C:\Users\djordje.subotic\Desktop\information_about_coping", DateTime.Now.ToString("dd.MM.yyy"), ".txt"));
            }
            catch(DirectoryNotFoundException de)
            {
                Console.WriteLine("There was an error.\nMessage:{0}\nTarget:{1}",de.Message, de.TargetSite);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("There was an error.\nMessage:{0}\nTarget:{1}", ane.Message, ane.TargetSite);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("There was an error.\nMessage:{0}\nTarget:{1}", ae.Message, ae.TargetSite);
            }


            string[] folder_with_pdf_or_word = null;
            string[] access_files = null;
            string access_file_with_path_name = null;
            string[] files_in_Directory = null;
            int? number_of_files_in_Directory = null;
            

            GetNumberOfFilesInDirectory(fromPath, ref folder_with_pdf_or_word, ref files_in_Directory, ref number_of_files_in_Directory);
            access_files = GetAccessFileFromPath(fromPath);
            if(access_files.Length > 0)
            {

                access_file_with_path_name = access_files[0];
            }

            sw.WriteLine("{0}**{1}**{2}", new String('#', 60), DateTime.Now.ToString(), new String('#', 60));

            sw.WriteLine("Method Execution Started time:**{0}**", DateTime.Now.ToString());
            sw.WriteLine("From Path:**{0}**", fromPath);
            sw.WriteLine("To Path:**{0}**",toPath);

            string cs = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source ={0};", access_file_with_path_name);
            object numOfRows = Migracija.GetNumberOfRowsInTable(cs, "RR");

            //Data For CSV Report ****************************************************************************************************************
            StreamWriter sw1 = null;
            try
            {
                sw1 = File.AppendText(String.Format("{0}_{1}.{2}", @"C:\Users\djordje.subotic\Desktop\data_exel", DateTime.Now.ToString("dd.MM.yyy"), ".txt"));
            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("There was an error.\nMessage:{0}\nTarget:{1}", de.Message, de.TargetSite);
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("There was an error.\nMessage:{0}\nTarget:{1}", ane.Message, ane.TargetSite);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("There was an error.\nMessage:{0}\nTarget:{1}", ae.Message, ae.TargetSite);
            }
            sw1.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", toPath.Split(new char[] { '_' })[1],political_folder_name_id, number_of_files_in_Directory, numOfRows, (int)numOfRows- number_of_files_in_Directory,DateTime.Now.ToString("dd.MM.yyy"));
            sw1.Close();
            
            int ditrict_id = Convert.ToInt32(political_folder_name_id);
            string district_name = toPath.Split(new char[] { '_' })[1];
            string connectionString = "Connection parameters replaced";
            int i = Migracija.SaveDataInformationToDatabase(connectionString, ditrict_id, district_name, (int)numOfRows, number_of_files_in_Directory);
            if (i == 1)
            {
                sw.WriteLine("Informations are saved to table Statistika.Number of saved rows is {0}", i);

            }
            else
            {
                sw.WriteLine("Nothing is saved to table Statistika");
            }
            /*************************************************************************************************************************************/

            Console.WriteLine("Number Of Rows is: **{0}**", numOfRows);
            sw.WriteLine("Number Of Files In Directory:**{0}**", number_of_files_in_Directory);
            sw.WriteLine("Number Of Rows In Access:**{0}**", numOfRows);

            if (folder_with_pdf_or_word.Length == 0)
            {
                sw.WriteLine("there is no Directory with word or pdf documents");
            }
            else
            {
                sw.WriteLine("Name Of Directory For Coping:**{0}**", folder_with_pdf_or_word[0]);
            }
           
            sw.WriteLine("Name Of Access File For Coping:**{0}**", access_file_with_path_name);

            Console.WriteLine(folder_with_pdf_or_word.Length);
            Console.WriteLine(access_file_with_path_name);
            Console.WriteLine(number_of_files_in_Directory);
            Console.WriteLine((int)numOfRows);

            if (folder_with_pdf_or_word.Length == 1 && access_file_with_path_name != null && number_of_files_in_Directory == (int)numOfRows)
            {
                sw.WriteLine("*****Number of files in directory and number of rows is equal, so we will start coping files and migration to central database.*****");
                string ac_n = GetAccessFileFromPathWithoutPath(access_file_with_path_name);
                Console.WriteLine("Acces name is {0}", ac_n);
                string access_file_from_previous_version = null;
                if (DoesonlyOneAccessFileExistInPath(toPath) == true)
                {
                    access_file_from_previous_version = (GetAccessFileFromPath(toPath))[0];
                    sw.WriteLine("Access file allready exists in To Path:**{0}**", access_file_from_previous_version);
                    DeleteExistingAccessFileFromPreviousVersion(access_file_from_previous_version);
                    Console.WriteLine("I am deleting access file now.Its Deleted");
                    sw.WriteLine("Access file:**{0}** is deleted.", access_file_from_previous_version);

                }

                Console.WriteLine("Copying newest access file.");
                sw.WriteLine($"Copy Access file from:**{fromPath}**.");
                sw.WriteLine($"Copy Access file to:**{toPath}**.");

                sw.WriteLine("Copy Access File Started time:**{0}**", DateTime.Now.ToString());
                string copied_access_file_full_name = CopyAccessFileToPoliticalDistrictFolder(access_file_with_path_name, toPath, ac_n);
                sw.WriteLine("Copy Access File Finished time:**{0}**", DateTime.Now.ToString());
                sw.WriteLine($"Access File is Copied.Copied file is in location: **{copied_access_file_full_name}**");
                sw.WriteLine("Copy Access File Finished time:**{0}**", DateTime.Now.ToString());

                //Connecting to this MS Access file and migrate
                Console.WriteLine("#####################################Start Database################################################################");
                string connectionParameters_access = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source ={0};", copied_access_file_full_name);
                string connectionParameters_SQL_Server = "Connection parameters replaced";

                sw.WriteLine("Access file connection:**{0}**", connectionParameters_access);
                sw.WriteLine("Sql Server Connection:**{0}**", connectionParameters_SQL_Server);
                int n = 0;
                sw.WriteLine("Starting Migration From Access To SQL Server.");
                sw.WriteLine("Database Migration Started Time:**{0}**", DateTime.Now.ToString());

                //Delete Existing Records in Database
                sw.WriteLine("Deleting existing records from database");
                int? nums_of_deleted_records = null;
                try
                {
                    nums_of_deleted_records = Migracija.DeleteAllRecordsForSpecifiedPoliticalDistrict(connectionParameters_SQL_Server, political_folder_name_id, "RR");
                    int nodr = nums_of_deleted_records ?? 0;
                    sw.WriteLine("Number of Deleted Rows is: {0}", nodr);
                }
                catch (Exception e)
                {

                    sw.WriteLine("Error occured while deleting existing records.");
                    sw.WriteLine("Error message.{0}", e.Message);
                }


                Migracija.MigrateDataFromOneDatabseToAnotherOne(connectionParameters_access, connectionParameters_SQL_Server, out n, true);

                sw.WriteLine("Database Migration Finished Time:**{0}**", DateTime.Now.ToString());
                sw.WriteLine("Ending Migration From Access To SQL Server.");
                sw.WriteLine("Number Of Migrated Rows:**{0}**", n);
                Console.WriteLine("###################################End Database###################################################################");

                string dir_for_files = toPath + "\\" + political_folder_name_id;
                Console.WriteLine("Directory: {0}", dir_for_files);
                sw.WriteLine("Directory For Copied Files:**{0}**", dir_for_files);

                bool dir_exists = Directory.Exists(dir_for_files);
                if (dir_exists == true)
                {
                    sw.WriteLine("Directory Allready Exist.");
                    sw.WriteLine("Deleting Directory....");
                    Directory.Delete(dir_for_files, true);
                    sw.WriteLine("Directory Deleted.");
                    sw.WriteLine("Copy Files Started Time:**{0}**", DateTime.Now.ToString());
                    sw.WriteLine("Creating new directory and copy files.");
                    CreateDirectoryAndCopyDocuments(dir_for_files, files_in_Directory);

                    sw.WriteLine("Copy Files Finished Time:**{0}**", DateTime.Now.ToString());
                    sw.WriteLine("Creating new directory and copy files.");


                }
                else
                {
                    sw.WriteLine("Directory does not exist.Creating new directory....");
                    CreateDirectoryAndCopyDocuments(dir_for_files, files_in_Directory);
                    sw.WriteLine("New Directory Created Successfuly.");
                }
                //Migrate Files To Database
                string[] all_files = Directory.GetFiles(dir_for_files);
                string connectionStringFiles = "Connection parameters replaced";/*"Data Source=DKP323-1;Database=Resenja;Integrated Security=SSPI";*/
                sw.WriteLine("Checking if there is files in database...");
                int? nums_of_deleted_records_files = Migracija.DeleteAllRecordsForSpecifiedPoliticalDistrict(connectionStringFiles, political_folder_name_id, "Fajlovi");
                int no = nums_of_deleted_records_files ?? 0;
                sw.WriteLine("Number of deleted files is {0}", no);
                sw.WriteLine("Starting inserting files to Database...");
                int n1 = Migracija.SaveFilesToDatabase(connectionStringFiles, political_folder_name_id, all_files, sw);
                sw.WriteLine("Number of inserted files is: {0}", n1);
                sw.WriteLine("Ending inserting files to Database...");

                /*StreamWriter sw1 = File.AppendText(@"C:\Users\djordje.subotic\Desktop\debili.txt");
                foreach (string s in Directory.GetFiles(dir_for_files))
                {
                    sw1.WriteLine("{0} - {1}",s, dir_for_files);
                }*/

            }
            else if (access_file_with_path_name != null)
            {
                /*###################################################################################################################################################*/
                sw.WriteLine("*****There is no directory with files but there is MS ACCESS database.Database migration will start even there is no directory with files*****");
                //string ac_n = GetAccessFileFromPathWithoutPath(access_file_with_path_name);
                //Console.WriteLine("Acces name is {0}", ac_n);
                /*string access_file_from_previous_version = null;
                if (DoesonlyOneAccessFileExistInPath(toPath) == true)
                {
                    access_file_from_previous_version = (GetAccessFileFromPath(toPath))[0];
                    sw.WriteLine("Access file allready exists in To Path:**{0}**", access_file_from_previous_version);
                    DeleteExistingAccessFileFromPreviousVersion(access_file_from_previous_version);
                    Console.WriteLine("I am deleting access file now.Its Deleted");
                    sw.WriteLine("Access file:**{0}** is deleted.", access_file_from_previous_version);

                }

                Console.WriteLine("Copying newest access file.");
                sw.WriteLine($"Copy Access file from:**{fromPath}**.");
                sw.WriteLine($"Copy Access file to:**{toPath}**.");

                sw.WriteLine("Copy Access File Started time:**{0}**", DateTime.Now.ToString());
                string copied_access_file_full_name = CopyAccessFileToPoliticalDistrictFolder(access_file_with_path_name, toPath, ac_n);
                sw.WriteLine("Copy Access File Finished time:**{0}**", DateTime.Now.ToString());
                sw.WriteLine($"Access File is Copied.Copied file is in location: **{copied_access_file_full_name}**");
                sw.WriteLine("Copy Access File Finished time:**{0}**", DateTime.Now.ToString());*/

                //Connecting to this MS Access file and migrate
                Console.WriteLine("#####################################Start Database################################################################");
                string connectionParameters_access = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source ={0};", access_file_with_path_name);
                string connectionParameters_SQL_Server = "Connection parameters replaced";

                sw.WriteLine("Access file connection:**{0}**", connectionParameters_access);
                sw.WriteLine("Sql Server Connection:**{0}**", connectionParameters_SQL_Server);
                int n = 0;
                sw.WriteLine("Starting Migration From Access To SQL Server.");
                sw.WriteLine("Database Migration Started Time:**{0}**", DateTime.Now.ToString());

                //Delete Existing Records in Database
                sw.WriteLine("Deleting existing records from database");
                int? nums_of_deleted_records = null;
                try
                {
                    nums_of_deleted_records = Migracija.DeleteAllRecordsForSpecifiedPoliticalDistrict(connectionParameters_SQL_Server, political_folder_name_id, "RR");
                    int nodr = nums_of_deleted_records ?? 0;
                    sw.WriteLine("Number of Deleted Rows is: {0}", nodr);
                }
                catch (Exception e)
                {

                    sw.WriteLine("Error occured while deleting existing records.");
                    sw.WriteLine("Error message.{0}", e.Message);
                }


                Migracija.MigrateDataFromOneDatabseToAnotherOne(connectionParameters_access, connectionParameters_SQL_Server, out n, false);

                sw.WriteLine("Database Migration Finished Time:**{0}**", DateTime.Now.ToString());
                sw.WriteLine("Ending Migration From Access To SQL Server.");
                sw.WriteLine("Number Of Migrated Rows:**{0}**", n);
                Console.WriteLine("###################################End Database###################################################################");

                /*#######################################################################################################################################*/

            }
            else
            {
                sw.WriteLine("Something went wrong.Number of files in folder is not equal to number of rows in database or access file is missing or directory with files is missing.");

            }

            sw.WriteLine("Method Execution Finished time:**{0}**", DateTime.Now.ToString());

            sw.Close();
        }                

        //Get all Folders From 4ZaKopiranjeURR
        public static IEnumerable<string> GetAllSubFolders(string folderName = @"\\srvfile1\Resenja o rusenju\4ZaKopiranjeuRR")
        {
            List<string> subFolders = new List<string>();
            try
            {
                subFolders = new List<string>();
                string[] allSubFolders = Directory.GetDirectories(folderName);
                foreach (string s in allSubFolders)
                {
                    subFolders.Add(s);
                }
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("An error occured.\n{0}", ae.Message);
            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("An error occured.\n{0}", de.Message);
            }
            return subFolders;
        } 

        //Match zip file name wit political district folder names 
        public static void MatchZipNameWithPoliticalDistrictFolderName(IEnumerable listOfZips, IEnumerable listOfFolders)
        {
           
            foreach (string i in listOfFolders)
            {
                Console.WriteLine("path in list of folders is: {0}", i);
                foreach (string j in listOfZips)
                {
                    Console.WriteLine("Path in zipfolders is {0}", j);
                    string filename = Path.GetFileName(j);
                    string id = filename.Split(new char[] { '_'})[0];
                    if (i.Split(Path.DirectorySeparatorChar).Last().StartsWith(id))
                    {
                        Console.WriteLine("There is match between i and. Start Unziping");

                        try
                        {
                            UnzipZipFile(j, i);

                        }
                        catch(System.IO.IOException w)
                        {
                            Console.WriteLine("MOJA GRESKA PANCEVO IO{0}", w.Message);
                        }
                    }
                    else
                    {
                        Console.WriteLine("There is no match");

                    }
                }
            }

        }

        //Unzip Content of Zip Files 
        public static void UnzipZipFile(string pathToZip, string pathToExtract)
        {
            ZipFile.ExtractToDirectory(pathToZip, pathToExtract);
        }

        //Read Data From CSV
        public static List<string> ReadFromCSV(string pathToCsv)
        {
            string[] data = null;
            List<string> datai = new List<string>();
            try
            {
                data = File.ReadAllLines(pathToCsv);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine(ae);
            }
            catch (FileNotFoundException fnf)
            {
                Console.WriteLine("There was an error.\n",fnf.Message);
            }

            foreach (string s in data)
            {
                datai.Add(s);
            }
            return datai;
        }

        public static List<string> ReadDataFromFolder(string path)
        {
            List<string> data =new List<string>();

            foreach (string h in Directory.EnumerateFiles(path))
            {
                data.Add(h.Split(Path.DirectorySeparatorChar).Last());
                
            }
            File.WriteAllLines(@"C:\Users\djordje.subotic\Desktop\SrkiFilenames.txt",Directory.EnumerateFiles(path));
            return data;
        }

        public static List<string> MatchCsvDataToFolder(List<string> csvNames, List<string> folderFileNames)
        {
           
            List<string> differentFiles = new List<string>();
            int inBoath = 0;
            foreach (string t in folderFileNames)
            {
                if (csvNames.Contains(t))
                {
                    Console.WriteLine("It Containes boath: {0}", t);
                    inBoath++;
                }
                else
                {
                    differentFiles.Add(t);
                    Console.WriteLine("Added to list of diferent files: {0}", t);
                }
            }

            Console.WriteLine("Number of files that are in boath folders is: {0}", inBoath);
            return differentFiles;

        }

        public static void SaveDiferencesToCsv(string path, List<string> data)
        {
            File.WriteAllLines(path, data);
        }
    }


}
