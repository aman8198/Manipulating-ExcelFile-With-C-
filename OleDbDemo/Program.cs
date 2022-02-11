using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OleDbDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\New folder\\One.xlsx; Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                oledbConn.Open();

                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet3$] where RegressionTest='no'", oledbConn);
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                DataSet ds = new DataSet();
                oleda.Fill(ds, "App");

                foreach (var m in ds.Tables[0].DefaultView)
                {
                    Console.WriteLine(((System.Data.DataRowView)m).Row.ItemArray[0] + " " + ((System.Data.DataRowView)m).Row.ItemArray[1] + " " + ((System.Data.DataRowView)m).Row.ItemArray[2] +" "+ ((System.Data.DataRowView)m).Row.ItemArray[3] +" "+ ((System.Data.DataRowView)m).Row.ItemArray[4]);

                }
                Console.ReadKey();

            }
            catch (Exception e)
            {
                Console.WriteLine("Error :" + e.Message);
                Console.Read();
            }
            finally
            {
                // Close connection
                oledbConn.Close();
            }
        }
    }
}
