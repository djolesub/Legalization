using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using System.Data.Common;
using System.IO;
using Npgsql;
using System.Runtime.Serialization.Formatters.Binary;

namespace Resenja_Folderi_Opstine
{
    class Migracija
    {
        public static DataSet GetAllDataFromTable(OleDbConnection conn)
        {
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            
            DataSet ds = new DataSet();
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = new OleDbCommand("SELECT * FROM RR", conn);
            try
            {
                da.Fill(ds, "RR");
                conn.Close();
            }
            catch (SystemException se)
            {
                Console.WriteLine("An error occured.\n", se.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            return ds;
        }

        public static byte[]  PreparingFileForSavingToDatabase(string filename, StreamWriter sw)
        {
            
            FileStream fs = null;
            byte[] buffer_file_content = null;
            try
            {
                fs = new FileStream(filename, FileMode.Open);
                buffer_file_content = new byte[fs.Length];

                try
                {
                    fs.Read(buffer_file_content, 0, (int)fs.Length);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error occured when reading content of file.\n{0}", e.Message);
                }
                finally
                {
                    fs.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error occured when creating file stream.\n{0}", e.Message);

            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
               
            }
            return buffer_file_content;
        }

        //Saving File Stream Content To Database
        public static void SaveFileToDatabase(byte[] buffered_content, string connectionString)
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(connectionString);

                SqlCommand sql = new SqlCommand("INSERT INTO RR_Files VALUES(@1)", conn);
                sql.Parameters.Add("@1", buffered_content);
                conn.Open();
                int i = sql.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("error While inserting file.\n{0}", e.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
           
        }




        public static void InsertDataToDatabase(DataSet ds, SqlConnection conn, bool right_case)
        {
            SqlDataAdapter da = new SqlDataAdapter();
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            try
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    da.InsertCommand = new SqlCommand("INSERT INTO RR([PO_id],[PO_naziv],[KO_id],[KO_naziv],[brojParcele],[podbrojParcele],[vrstaRR],[izvezen],[dokument],[redniBroj],[originalNaziv],[napomena],[datumKreiranja],[FajlDaNe]) VALUES (@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14)", conn);

                    da.InsertCommand.Parameters.Add("@1", dr[1]);
                    da.InsertCommand.Parameters.Add("@2", dr[2]);
                    da.InsertCommand.Parameters.Add("@3", dr[3]);
                    da.InsertCommand.Parameters.Add("@4", dr[4]);
                    da.InsertCommand.Parameters.Add("@5", dr[5]);
                    da.InsertCommand.Parameters.Add("@6", dr[6]);
                    da.InsertCommand.Parameters.Add("@7", dr[7]);
                    da.InsertCommand.Parameters.Add("@8", dr[8]);
                    da.InsertCommand.Parameters.Add("@9", dr[9]);
                    da.InsertCommand.Parameters.Add("@10", dr[10]);
                    da.InsertCommand.Parameters.Add("@11", dr[11]);
                    da.InsertCommand.Parameters.Add("@12", dr[12]);
                    da.InsertCommand.Parameters.Add("@13", dr[13].ToString());
                    da.InsertCommand.Parameters.Add("@14", right_case);


                    da.InsertCommand.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception While Inserting To central Db.\n{0}",e.Message);
                Console.WriteLine("Method That Throws Exception: {0}", e.TargetSite);
                Console.WriteLine("Stack trace: {0}", e.StackTrace);


            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }

        }

        public static object GetNumberOfRowsInTable(string connection, string tableName)
        {
            OleDbConnection conn = null;
            var number_of_rows = (object)null;
            try
            {
                conn = new OleDbConnection(connection);
                OleDbCommand sql = new OleDbCommand(String.Format("SELECT Count(*) FROM {0}", tableName), conn);
                conn.Open();
                number_of_rows = (int)sql.ExecuteScalar();
            }
            catch (InvalidOperationException ioe)
            {
                Console.WriteLine("There was an error while connecting.\n{0}", ioe.Message);
            }
            catch (OleDbException odbe)
            {
                Console.WriteLine("There was an error while connecting.\n{0}", odbe.Message);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
            return number_of_rows;
            
        }

        //Delete all records for specified political district from Database
        public static int? DeleteAllRecordsForSpecifiedPoliticalDistrict(string connectionString1, string political_district_id, string table)
        {
            SqlConnection conn = new SqlConnection(connectionString1);
            SqlCommand sql = new SqlCommand($"DELETE FROM {table} WHERE PO_id={political_district_id}", conn);
            int? numberOfDeletedRows = null;
            try
            {
                conn.Open();
                try
                {
                    numberOfDeletedRows = sql.ExecuteNonQuery();
                }
                catch (InvalidOperationException ioe)
                {
                    Console.WriteLine("There was an errro.\n{0}", ioe.Message);
                }
                catch (SqlException se)
                {
                    Console.WriteLine("There was an error.\n{0}", se.Message);
                }
                conn.Close();
            }
            catch (InvalidOperationException ioe)
            {
                Console.WriteLine("There was an error.\n{0}", ioe.Message);
            }
            catch (SqlException se)
            {
                Console.WriteLine("There was an error.\n{0}", se.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            return numberOfDeletedRows;

        }

        public static void MigrateDataFromOneDatabseToAnotherOne(string connectionString1, string connectionString2, out int number_of_migrated_rows, bool right_case, string migrated_rows_file_name = @"C:\Users\djordje.subotic\Desktop\nums.txt")
        {

            OleDbConnection conn1 = new OleDbConnection(connectionString1);
            SqlConnection conn2 = new SqlConnection(connectionString2);
            DataSet data_to_migrate = GetAllDataFromTable(conn1);
            number_of_migrated_rows = data_to_migrate.Tables[0].Rows.Count;
            InsertDataToDatabase(data_to_migrate, conn2, right_case);
            StreamWriter migrated_rows = null;
            try
            {
                migrated_rows = File.AppendText(migrated_rows_file_name);

                migrated_rows.WriteLine(String.Format("{0}{1}{2}", new String('#', 20), DateTime.Now.ToString(), new String('#', 20)));
                migrated_rows.WriteLine("Number of migrated rows is:\t{0}", number_of_migrated_rows);
                migrated_rows.Close();
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
                if (migrated_rows != null)
                {
                    migrated_rows.Close();
                }
            }


        }

        //Saving basic Satatistics to database Statistika
        public static int SaveDataInformationToDatabase(string connectionString, int ditrict_id, string district_name, int access_rows, int? number_of_files)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommand sql = new SqlCommand("INSERT INTO Statistika VALUES(@1,@2,@3,@4, @5, @6)", conn);
            int number_of_inserted_rows = 0;
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            try
            {

                da.InsertCommand = sql;
                da.InsertCommand.Parameters.Add("@1", ditrict_id);
                da.InsertCommand.Parameters.Add("@2", district_name);
                da.InsertCommand.Parameters.Add("@3", access_rows);
                int nof = number_of_files ?? 0;
                da.InsertCommand.Parameters.Add("@4", nof);
                da.InsertCommand.Parameters.Add("@5", Math.Abs(access_rows - nof));
                da.InsertCommand.Parameters.Add("@6", DateTime.Now.ToString("dd.MM.yyy"));

                number_of_inserted_rows  = da.InsertCommand.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occured.{0}", e.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            return number_of_inserted_rows;


        }
        //Saving File Paths To Database
        public static int SaveFilesToDatabase(string connectionString,string id, string[] paths, StreamWriter sw)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommand sql = null; 
            int number_of_inserted_rows_files = 0;
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            try
            {
                foreach (string path in paths)
                {
                    string name = path.Split(new char[] { '\\' })[path.Split(new char[] { '\\' }).Length - 1];
                    string folder_name_id = path.Split(new char[] { '\\' })[path.Split(new char[] { '\\' }).Length - 2];
                    string folder_name = path.Split(new char[] { '\\' })[path.Split(new char[] { '\\' }).Length - 3];
                    string server_folder_name = String.Format("{0}\\{1}\\{2}\\{3}","E:\\Resenja o rusenju",folder_name, folder_name_id,name);
                    sql = new SqlCommand("INSERT INTO Fajlovi VALUES(@1,@2,@3)", conn);
                    da.InsertCommand = sql;
                    da.InsertCommand.Parameters.Add("@1", server_folder_name);
                    da.InsertCommand.Parameters.Add("@2", name);
                    da.InsertCommand.Parameters.Add("@3", id);
                    number_of_inserted_rows_files += da.InsertCommand.ExecuteNonQuery();
                }

                
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occured.{0}", e.Message);
                sw.WriteLine("An error occured while saving files to database.");
                sw.WriteLine("Error message: {0}", e.Message);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            return number_of_inserted_rows_files;


        }

        //Get all KO_IDs from RR table in Resenja DB
        public static IEnumerable<int> GetAllKOIds(string connection)
        {
            List<int> KO_IDs= new List<int>();
            SqlConnection conn = new SqlConnection(connection);
            try
            {
                conn.Open();
                SqlCommand sql = new SqlCommand(@"SELECT KO_id FROM RR", conn);
                SqlDataReader reader = sql.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        KO_IDs.Add(Convert.ToInt32(reader["KO_id"]));
                    }
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            return KO_IDs.Distinct<int>();
        }

        //Connect To Posgis NIGP Database
        public static NpgsqlConnection NigpConnection(string connection)
        {
            NpgsqlConnection conn = new NpgsqlConnection(connection);
            try
            {
                conn.Open();
                Console.WriteLine("Successfuly connected to {0} version {1}", conn.DataSource, conn.PostgreSqlVersion);
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();             
            }
            return conn;
        }

        //Using BulkCopy To Migrate NIGP Geometry to Sql Server table GeometrijaNIGPParcele
        public static void MigrateGeometry(string connectionNIGP, string connectionSqlServer,  IEnumerable<int> KO_IDs)
        {
            
            NpgsqlConnection nigp_connection = new NpgsqlConnection(connectionNIGP);
            SqlConnection sql_server_connection = new SqlConnection(connectionSqlServer);
            DataTable table_from_nigp = new DataTable();
            try
            {
                nigp_connection.Open();
                foreach (int ko_id in KO_IDs)
                {
                    Console.WriteLine("Starting For KO: {0}",ko_id);
                    string sql = String.Format(@"SELECT names of columns
                                                    FROM table
                                                    where maticnibrojko = {0}", ko_id);
                    NpgsqlCommand sql_nigp = new NpgsqlCommand(sql, nigp_connection);
                    NpgsqlDataAdapter sql_server_adapter = new NpgsqlDataAdapter(sql_nigp);
                    sql_server_adapter.Fill(table_from_nigp);
                    Console.WriteLine("Number of records is: {0}", table_from_nigp.Rows.Count);
                    try
                    {
                        sql_server_connection.Open();
                        int a = InsertGeometryFromNIGP(table_from_nigp, sql_server_connection);
                        Console.WriteLine("Finishing For KO: {0}", ko_id);
                        Console.WriteLine("Number of migrated records is: {0}", a);
                        table_from_nigp.Clear();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error {0}", e.Message);
                    }
                    finally
                    {
                        sql_server_connection.Close();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception\n{0}", e.Message);
            }

            finally
            {
                if (sql_server_connection.State == ConnectionState.Open)
                    sql_server_connection.Close();
            }
        }

        //Inserting geometry From NIGP_ To SQL Server 
        public static int InsertGeometryFromNIGP(DataTable ds, SqlConnection conn)
        {
            SqlDataAdapter da = new SqlDataAdapter();
            int migrated_rows = 0;
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            try
            {
                foreach (DataRow dr in ds.Rows)
                {
                    da.InsertCommand = new SqlCommand("INSERT INTO GeometrijaNIGPparcele(names of columns in parcela table) VALUES (@1,@2,@3,@4,@5,@6,@7)", conn);

                    da.InsertCommand.Parameters.Add("@1", dr[0]);
                    da.InsertCommand.Parameters.Add("@2", dr[1]);
                    da.InsertCommand.Parameters.Add("@3", dr[2]);
                    da.InsertCommand.Parameters.Add("@4", dr[3]);
                    da.InsertCommand.Parameters.Add("@5", dr[4]);
                    da.InsertCommand.Parameters.Add("@6", dr[5]);
                    da.InsertCommand.Parameters.Add("@7", dr[6]);
                    
                    migrated_rows = da.InsertCommand.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception While Inserting To central Db.\n{0}", e.Message);
                Console.WriteLine("Method That Throws Exception: {0}", e.TargetSite);
                Console.WriteLine("Stack trace: {0}", e.StackTrace);


            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            return migrated_rows;

        }
    }
}

