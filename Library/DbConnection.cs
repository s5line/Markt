using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace Library
{
    public class DbConnection
    {
        private static SqlConnection sqlConnection;
        private static string connectionString = @"Data Source=c:\Development\Markt\Markt\Markt\DBMarkt.sqlite;Version=3";
        

        public static void sqlQeury(string query)
        {
            try
            {
                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();

                SqlCommand cmd = new SqlCommand(query, sqlConnection);
                cmd.ExecuteNonQuery();

                cmd.Dispose();
                sqlConnection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataSet sqlGetDataSet(string query)
        {
            DataSet dataSet = new DataSet();



            return dataSet;
        }

    }
}
