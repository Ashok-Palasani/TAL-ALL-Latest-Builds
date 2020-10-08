using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data.Common;
using System.Data;
using System.Collections.ObjectModel;
using System.Collections;
using System.Collections.Specialized;
using System.Data.SqlClient;
using System.Configuration;

namespace AutoReportMail
{
    class MsqlConnection : IDisposable
    {
        //static String ServerName = "TCP:10.20.10.65,1433";
        //static String username = "sa";
        //static String password = "srks4$tsal";
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

        //public MySqlConnection msqlConnection = new MySqlConnection("server = " + ServerName + ";userid = " + username + ";Password = " + password + ";database = " + DB + ";port = " + port + ";persist security info=False");
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

        void IDisposable.Dispose()
        { }

    }
}
