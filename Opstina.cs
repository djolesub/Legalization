using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace RO_Functions
{
    public  class Opstina
    {
        //Defining fields
        private FileInfo accessDB;
        private DirectoryInfo all_files;
        private string topFolder;
        private string opstina_id;
        private OleDbConnection conn;

        //Defining properties
        public FileInfo AccessDB { get { return accessDB; } set { accessDB = value; }}
        public DirectoryInfo AllFiles { get { return all_files; } set { all_files = value; } }
        public string TopFolder { get { return topFolder; } set { topFolder = value; } }
        public string OpstinaID { get { return opstina_id; } set { opstina_id = value; } }
        public OleDbConnection Connection { get { return conn; } set { conn = value; } }

        //Defining Constructor
        public Opstina(string topFolder)
        {
            string access_db_path = Path.Combine(topFolder, "UMP_RR.accdb");
            string files_path = Path.Combine(topFolder,Helper.GetFolderIDFromFolderPath(topFolder));

            AccessDB = new FileInfo(access_db_path);
            AllFiles = new DirectoryInfo(files_path);
            TopFolder = topFolder;
            OpstinaID = Helper.GetFolderIDFromFolderPath(topFolder);
        }

        //Setting connection to access file
        public void SetConnectionParams()
        {
            this.Connection = new OleDbConnection(String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source ={0};", this.AccessDB));
        }

        //Open Connection
        public void OpenConnection()
        {
            try
            {
                if (Connection.State == ConnectionState.Closed)
                    Connection.Open();
            }
            catch (InvalidOperationException ie)
            {
                Console.WriteLine("There was an error while connecting.\n{0}", ie.Message);
            }
            catch (OleDbException ode)
            {
                Console.WriteLine("There was an error while connecting.\n{0}", ode.Message);
            }
        }

        //Close Connection
        public void CloseConnection()
        {
            if (Connection.State == ConnectionState.Open)
                Connection.Close();    
        }

        //Get NUmber of rows in access file
        public int GetNumberOfRowsInAccessFile()
        {
            int rowsReturned = 0;
            try
            {
                this.OpenConnection();
                string query = @"SELECT COUNT(ID) FROM RR";
                OleDbCommand sql = new OleDbCommand(query, Connection);
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
                this.CloseConnection();
            }

            return rowsReturned;
        }

        //Get all doc names from database
        public IEnumerable<string> GetDocumentsFromDatabase()
        {
            IList<string> documents = new List<string>();
            try
            {
                this.OpenConnection();
                string query = @"SELECT dokument FROM RR";
                OleDbCommand sql = new OleDbCommand(query, Connection);
                OleDbDataReader reader = sql.ExecuteReader();
                if(reader.HasRows)
                {
                    while (reader.Read())
                    {
                        documents.Add(reader[0].ToString()/*.Split(new char[] { '.' })[0]*/);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
            }
            finally
            {
                this.CloseConnection();
            }
            return documents;
        }

        //Delete From Database 

        //Print documents names from database
        public void PrintAllDocumentNames()
        {
            this.GetDocumentsFromDatabase().ToList().ForEach( docname => 
                                                                Console.WriteLine(docname)
                                                            );
        }

        //Get document ids
        public IEnumerable<string> GetAllFiles()
        {
            IEnumerable<string> files = AllFiles.EnumerateFiles().Select(fileinfo => fileinfo.Name/*.Split(new char[] { '.'})[0]*/);
            return files;
        }

        //Print All document names
        public void PrintAllfiles()
        {
            AllFiles.EnumerateFiles().ToList().ForEach( fileinfo => 
                                                            { Console.WriteLine(fileinfo.Name.Split(new char[] { '.' })[0]);
                                                     });
        }

        //Check if all elements in collection are unique
        public Boolean IsCollectionUnique(IEnumerable<string> collection)
        {
            HashSet<string> unique = new HashSet<string>(collection); 
            return collection.Count() == unique.Count();
        }

        //Get Duplicates
        public IEnumerable<string> GetDuplicates(IEnumerable<string> collection)
        {
            IEnumerable<string> duplicateItems = collection.GroupBy(x => x)
                                                            .Where(x => x.Count() > 1)
                                                            .Select(x => x.Key);
            return duplicateItems;
        }



        //Get number of files in folder 
        public int GetNumberOfFiles()
        {
            int num = 0;
            try
            {
               num = AllFiles.EnumerateFiles().Where(fileinfo => fileinfo.Name.EndsWith(".pdf") ||
                                                                 fileinfo.Name.EndsWith(".doc") ||
                                                                 fileinfo.Name.EndsWith(".docx")
                                                                 ).Count();
            }
            catch (DirectoryNotFoundException de)
            {
                Console.WriteLine("Error: {0}", de.Message);
            }
            return num;    
        }

        public IEnumerable<string> OnlyInFile()
        {
            IEnumerable<string> filenames = this.GetAllFiles();
            IEnumerable<string> docnames = this.GetDocumentsFromDatabase();

            return  filenames.Except<string>(docnames);
        }

        public IEnumerable<string> OnlyInDatabase()
        {
            IEnumerable<string> filenames = this.GetAllFiles();
            IEnumerable<string> docnames = this.GetDocumentsFromDatabase();

            return docnames.Except<string>(filenames);
        }

        public IEnumerable<string> InFileAndDatabase()
        {
            IEnumerable<string> filenames = this.GetAllFiles();
            IEnumerable<string> docnames = this.GetDocumentsFromDatabase();
            return docnames.Intersect<string>(filenames);
        }


    }
}
