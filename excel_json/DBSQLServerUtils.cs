using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_json
{
    class DBSQLServerUtils
    {
        public static SqlConnection
            GetDBConnection(string datasource, string database)
        {
            //string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
            string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + "; Connection Timeout=45; Integrated Security = True;";          

            SqlConnection conn = new SqlConnection(connString);

            return conn;
        }
    }
}
