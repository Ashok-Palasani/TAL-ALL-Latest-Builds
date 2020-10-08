using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MimicsUpdation 
{
    class MsqlConnection : IDisposable
    {
        //Server
        //static String ServerName = @"SIEMENS\SQLEXPRESS";
        //static String username = "sa";
        //static String password = "srks4$taml";
        //static String port = "3306";
        //static String DB = "i_facility_taml";


        public static String ServerName = @"" + ConfigurationManager.AppSettings["ServerName"]; //SIEMENS\SQLEXPRESS
        public static String username = ConfigurationManager.AppSettings["username"]; //sa
                                                                                      //static String password = "srks4$";//server
        public static String password = ConfigurationManager.AppSettings["password"];
        public static String port = "3306";
        public static String DB = ConfigurationManager.AppSettings["DB"];// i_facility_tsal //Common
        public static String Schema = ConfigurationManager.AppSettings["Schema"];  //Schema Name

        //Tata Server
        //static String ServerName = @"SRKSDEV001-PC\SQLSERVER17"; 
        //static String username = "sa";
        //static String password = "srks4$";
        //static String DB = "i_facility_taml";


        //Local
        //static String ServerName = "localhost";
        //static String username = "root";
        //static String password = "srks4$";
        //static String port = "3306";
        //static String DB = "i_facility";

        public SqlConnection msqlConnection = new SqlConnection(@"Data Source = " + ServerName + ";User ID = " + username + ";Password = " + password + ";Initial Catalog = " + DB + ";Persist Security Info=True");

        public void open()
        {
            if (msqlConnection.State != System.Data.ConnectionState.Open)
                msqlConnection.Open();
        }

        public void close()
        {
            msqlConnection.Close();
        }

        public void Dispose()
        {
            msqlConnection.Dispose();
            GC.SuppressFinalize(this);
        }

        void IDisposable.Dispose()
        {
        }
    }
}
