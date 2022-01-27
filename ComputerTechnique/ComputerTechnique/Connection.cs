using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ComputerTechnique
{
    class Connection
    {
        public static SqlConnection sqlConnection = new SqlConnection(@"Data Source=DESKTOP-HIKNFP1\SQLEXPRESS; Initial Catalog=ComputerTechnique; Integrated Security=SSPI");

        public static bool connectOpen()
        {
            bool temp = false;
            try
            {
                sqlConnection.Open();
                temp = true;
            }
            catch
            {
                temp = false;
            }
            return temp;
        }

        public static void connectClose()
        {
            try
            {
                sqlConnection.Close();
            }
            catch { }
        }
    }
}
