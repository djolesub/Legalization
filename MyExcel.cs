using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace RO_Functions
{
    class MyExcel
    {
        //UpdateOrInsert metod
        public static void UpdateOrInsert(OleDbConnection conn, int maticni, int numFolder, int numBaza, string sheetName, StreamWriter sw)
        {
            IList<string> kos = GetMatBrIDs(conn, sheetName);
            DateTime s = DateTime.Now;
            string vreme = String.Format(@"{0}.{1}", s.Day, s.Month);

            if (kos.Contains<string>(maticni.ToString()) == true)
            {
                try
                {
                    sw.WriteLine("{0}Calling Method: {1}", Helper.Indent(10), String.Format("UpdateRowCells(OpstinaID={0}, FilesInFolder={1}, RowsInAccess={2}, Time={3}, SheetName={4})", maticni, numFolder, numBaza, vreme, sheetName));
                    UpdateRowCells(conn, maticni, numFolder, numBaza, vreme, sheetName);  
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception Message: {0}.\n\n{1}", e.Message, e.StackTrace);
                    sw.WriteLine("{0}Error occured: {1}", Helper.Indent(10), e.Message);
                    sw.WriteLine("{0}Method that throws error: {1}", Helper.Indent(10), e.TargetSite);
                    sw.Close();
                }
                finally
                {
                    
                }
            }
            else
            {
                try
                {
                    sw.WriteLine("{0}Calling Method: {1}", Helper.Indent(10), String.Format("InsertNewRow(OpstinaID={0}, FilesInFolder={1}, RowsInAccess={2}, Time={3}, SheetName={4})", maticni, numFolder, numBaza, vreme, sheetName));
                    InsertNewRow(conn, maticni, numFolder, numBaza, vreme, sheetName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception Message: {0}.\n\n{1}", e.Message, e.StackTrace);
                    sw.WriteLine("{0}Error occured: {1}", Helper.Indent(10), e.Message);
                    sw.WriteLine("{0}Method that throws error: {1}", Helper.Indent(10), e.TargetSite);
                    sw.Close();
                }
                finally
                {
                   
                }
            }
            sw.Close();
        }

        //Update existing row in excel file
        public static int UpdateRowCells(OleDbConnection conn, int maticni, int numFolder, int numBaza, string datum, string sheetName)
        {
            int error_type = Helper.GetErrorType(numFolder, numBaza);
            Console.WriteLine("Error type: {0}", error_type);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            string sqlCommand = String.Format(@"UPDATE [{0}$] SET Folder={1}, Baza={2}, Datum={3}, TipGreske={4} WHERE MaticniBR={5};", sheetName, numFolder, numBaza, datum, error_type, maticni);
            Console.WriteLine("Update Row:{0}", sqlCommand);
            OleDbCommand sql = new OleDbCommand(sqlCommand, conn);
            int rowsAffected = sql.ExecuteNonQuery();
            Console.WriteLine("Updated rows: {0}", rowsAffected);
            if (conn.State == ConnectionState.Open)
                conn.Close();
            return rowsAffected;
        }

        //Insert new row into exel file
        public static int InsertNewRow(OleDbConnection conn, int maticni, int numFolder, int numBaza, string datum, string sheetName)
        {
            int error_type = Helper.GetErrorType(numFolder, numBaza);
            Console.WriteLine("Error type: {0}", error_type);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            string sqlCommand = String.Format(@"INSERT INTO [{0}$] (MaticniBR, Folder, Baza, TipGreske, Datum) VALUES ({1},{2},{3},{4}, {5})", sheetName, maticni, numFolder, numBaza, error_type, datum);
            OleDbCommand sql = new OleDbCommand(sqlCommand, conn);
            int insertedRows = sql.ExecuteNonQuery();
            Console.WriteLine("Number of inserted rows: {0}", insertedRows);
            if (conn.State == ConnectionState.Open)
                conn.Close();
            return insertedRows;
        }

        //Get List of all MatBR from Excel file
        public static IList<string> GetMatBrIDs(OleDbConnection conn, string sheetName)
        {
            List<string> kos = new List<string>();
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            string sqlCommand = String.Format(@"SELECT MaticniBR FROM [{0}$]", sheetName);

            OleDbCommand sql1 = new OleDbCommand(sqlCommand, conn);
            OleDbDataReader reader = sql1.ExecuteReader();
            Boolean exist = false;
            while (reader.Read())
            {
                if (reader.HasRows)
                {
                    //Console.WriteLine(reader[0].ToString());
                    kos.Add(reader[0].ToString());
                }
            }
            if (conn.State == ConnectionState.Open)
                conn.Close();
            return kos;
        }
    }
}
