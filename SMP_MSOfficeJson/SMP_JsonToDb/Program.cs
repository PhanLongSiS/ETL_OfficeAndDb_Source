using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMP
{
    public class Program
    {
        static string connectionString = @"Data Source=snap.techlinkvn.com,1444;Initial Catalog=OffficeETL;Persist Security Info=True;User ID=officeetl;Password=9dj929djwjw064jfw";
        static SqlConnection sqlCon = null;
        static int Main(string[] args)
        {
            string fileName = args[0];
            string json = File.ReadAllText(fileName);
            //Getdata
            string machineName = Environment.MachineName;
            DateTime createdDate= DateTime.Now;
            string windowsUser= Environment.UserName;

            //Connect to Db
            if (sqlCon==null)
            {
                sqlCon = new SqlConnection(connectionString);
            }
            if (sqlCon.State==ConnectionState.Closed)
            {
                sqlCon.Open();
            }
            //
            SqlCommand cmd = new SqlCommand("insert into formPDHL values ('"+json+"','"+createdDate+"','"+machineName+"','"+windowsUser+"','','')",sqlCon);    
            int respone=cmd.ExecuteNonQuery();
            if (respone>0)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }
    }
}
