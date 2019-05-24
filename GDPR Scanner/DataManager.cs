using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace GDPR_Scanner
{
    class DataManager
    {
        public SqlConnection getStatsConnection()
        {
            string Server = ConfigurationManager.AppSettings["DB_Path"];
            string Username = ConfigurationManager.AppSettings["DB_User"];
            string Password = ConfigurationManager.AppSettings["DB_Password"];
            string Database = ConfigurationManager.AppSettings["DB_Database"];
            string connectionString = "Data Source=" + Server + "; User ID=" + Username + "; Password=" + Password + ";MultipleActiveResultSets=True; Initial Catalog=" + Database;
            SqlConnection SQLCon = new SqlConnection();
            try
            {
                SQLCon.ConnectionString = connectionString;
            }
            catch { }
            return SQLCon;
        }
    }
}
