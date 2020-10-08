using System;
using System.Collections.Generic;
using System.Text;
using System.Data.Common;
using System.Data;
using System.Collections.ObjectModel;
using System.Collections;
using System.Collections.Specialized;
using System.Data.SqlClient;
using System.Configuration;

namespace TataMySqlConnection
{
    class MsqlConnection:IDisposable
    {
        //static String ServerName = @"SRKSDEV001-PC\SQLSERVER17";
        //static String username = "sa";
        //static String password = "srks4$";
        //static String port = "3306";
        //static String DB = "i_facility_tsal";

       public static String ServerName = ConfigurationManager.AppSettings["ServerName"];
        public static String username = ConfigurationManager.AppSettings["username"];
        public static String password = ConfigurationManager.AppSettings["password"];
        public static String port = "3306";
        public static String DB = ConfigurationManager.AppSettings["DB"];
        public static String DbName = ConfigurationManager.AppSettings["DbName"];
        public static String Host = ConfigurationManager.AppSettings["host"];
        public static int Portmail = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]);

        public static String UserNamemail = ConfigurationManager.AppSettings["UserNameemail"];
        public static String Passwordmail = ConfigurationManager.AppSettings["Passwordemail"];
        public static String Domain = ConfigurationManager.AppSettings["Domain"];
        public static String IsSSL = ConfigurationManager.AppSettings["SSLENABLE"];

        public SqlConnection msqlConnection = new SqlConnection(@"Data Source = " + ServerName + ";User ID = " + username + ";Password = " + password + ";Initial Catalog = " + DB + ";Persist Security Info=True");

        public void open()
        {
            if (msqlConnection.State != System.Data.ConnectionState.Open)
                msqlConnection.Open();
        }

        public void close()
        {
            if (msqlConnection.State != System.Data.ConnectionState.Closed)
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
