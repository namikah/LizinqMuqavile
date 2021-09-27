using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Reflection;

namespace Nsoft
{
    class MyData
    {
        public static OleDbDataAdapter oledbadapter1;
        public static OleDbConnection oledbconnection1;
        public static DataTable dtmainUnvanlar;
        public static DataTable dtmainSurucuMelumat;
        public static DataTable dtmainNomre;
        public static DataTable dtmain;
        //public static DataTable dtmainNomreAxtar;
        public static DataTable dtmainParol;
        public static DataTable dtmainEtibarbame;
        public static DataTable dtmainEtibarnameArxiv;
        public static DataTable dtmainSurucu;
        public static DataTable dtmainMuqavileRekvizit;
        public static DataTable dtmainSuruculer;
        public static DataTable dtmainArxiv;
        public static DataTable dtmainEtibarnameNomre;
        //public static DataTable dtmainn;
        //public static DataTable dtmainOdenisler;
        public static DataTable dtmainQeydler;
        public static DataTable dtmainLisenziya;
        public static DataTable dtmainMMX;
        public static DataTable dtmainsozverenler;
        public static DataTable dtmainDYPUmumiMelumat;
        public static DataTable dtmaintelefon;
        public static DataTable dtmainemekhaqqi;
        public static DataTable dtmainrekvizitler;
        public static DataTable dtmainplateshkanomre;
        public static DataTable dtmainedv;
        public static DataTable dtmainpEDVnomre;
        public static DataTable dtmainTeyinat;

        public static void CreateConnection(string bazaName)
        {
            oledbconnection1 = new OleDbConnection();
            oledbadapter1 = new OleDbDataAdapter();
            oledbconnection1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + bazaName + "'";
        }    //elaqe yaratmaq

        public static void selectCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.SelectCommand = new OleDbCommand();
            oledbadapter1.SelectCommand.Connection = oledbconnection1;
            oledbadapter1.SelectCommand.CommandText = commandText;
        }

        public static void insertCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.InsertCommand = new OleDbCommand();
            oledbadapter1.InsertCommand.Connection = oledbconnection1;
            oledbconnection1.Open();
            oledbadapter1.InsertCommand.CommandText = commandText;
            oledbadapter1.InsertCommand.ExecuteNonQuery();
            oledbconnection1.Close();
        }
        public static void deleteCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.DeleteCommand = new OleDbCommand();
            oledbadapter1.DeleteCommand.Connection = oledbconnection1;
            oledbconnection1.Open();
            oledbadapter1.DeleteCommand.CommandText = commandText;
            oledbadapter1.DeleteCommand.ExecuteNonQuery();
            oledbconnection1.Close();
        }

        public static void updateCommand(String bazaName, String commandText)
        {
            CreateConnection(bazaName);
            oledbadapter1.UpdateCommand = new OleDbCommand();
            oledbadapter1.UpdateCommand.Connection = oledbconnection1;
            oledbconnection1.Open();
            oledbadapter1.UpdateCommand.CommandText = commandText;
            oledbadapter1.UpdateCommand.ExecuteNonQuery();
            oledbconnection1.Close();
        }

        public static string appInfo()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string result = "Company Name: " + fvi.CompanyName
                + Environment.NewLine + "Product Name: " + fvi.ProductName
                + Environment.NewLine + "File Location : " + fvi.FileName
                + Environment.NewLine + "Product Version: " + fvi.ProductVersion
                + Environment.NewLine
                + Environment.NewLine + "Comments: " + fvi.Comments
                + Environment.NewLine
                + Environment.NewLine + fvi.LegalCopyright;

            return result;
        }

    }
}
