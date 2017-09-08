using System;
using System.Collections.Generic;

using System.Web;
using System.Data;
using System.Data.OleDb;
/// <summary>
/// Excel 的摘要说明
/// </summary>
namespace Excel2Json
{
    public class Excel
    {

        public static DataSet SelectFromXLS(string file, string sheet)
        {
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Extended Properties='Excel 8.0;IMEX=1;HDR=Yes'"; ;
            OleDbConnection conn = new OleDbConnection(connStr);
            OleDbDataAdapter da = null;
            DataSet ds = new DataSet();
            try
            {
                conn.Open();
                da = new OleDbDataAdapter("select * from [" + sheet + "]", conn);
                da.Fill(ds, "SelectResult");
            }
            catch (Exception e)
            {
                conn.Close();
                throw e;
            }
            finally
            {
                conn.Close();
            }
            return ds;
        }

        public static DataTable GetXlsSheetName(string file)
        {
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Extended Properties='Excel 8.0;IMEX=1;HDR=No'"; ;
            OleDbConnection conn = new OleDbConnection(connStr);
            //OleDbDataAdapter da = null;
            DataTable ExcelSheets = new DataTable();
            try
            {
                conn.Open();
                ExcelSheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            }
            catch (Exception e)
            {
                conn.Close();
                throw e;
            }
            finally
            {
                conn.Close();
            }
            return ExcelSheets;
        }



        public static string ConvertToSQLSheetName(string SheetName)
        {
            return "[" + SheetName + "]";
        }
    }
}