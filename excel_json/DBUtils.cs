using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_json
{
    class DBUtils
    {
        public static SqlConnection GetDBConnection()
        {
            string datasource = @"ATSVR12";
            string database = "SMPO";
        //    string username = "sa";
        //    string password = "Killsews777";

            return DBSQLServerUtils.GetDBConnection(datasource, database);
        }
    }
}
