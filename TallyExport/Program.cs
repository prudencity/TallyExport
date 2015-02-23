using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using Microsoft.Win32;
using System.IO;
using System.Management;
using System.Reflection;

namespace TallyExport
{
    static class Program
    {

        public static string connectionString = null;

        public static string DataBaseName = "";
        public static string ServerName = "";
        public static string ServerUserId = "";
        public static string ServerPassword = "";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            class_DataAccess ObjDataAccess = new class_DataAccess();

            bool checkFile = false;

            checkFile = System.IO.File.Exists("Finance.stl");

            if (checkFile == false)
            {
                MessageBox.Show("Error in Connection to Database. Contact Admin.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                Application.Exit();
                return;
            }
            else
            {
                StreamReader objReader = new StreamReader("Finance.stl");
                Program.connectionString = objReader.ReadLine();
                objReader.Close();
                objReader.Dispose();

                try
                {
                    ObjDataAccess.CreateConnection();
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Error in Connection to Database. Contact Admin", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }

            string[] connSplit = Program.connectionString.Split(';');
            if (connSplit.Length == 6)
            {
                string[] DataSourceName = connSplit[0].Split('=');
                string[] DataBaseName = connSplit[1].Split('=');
                Program.DataBaseName = DataBaseName[1].ToString();
            }

            else if (connSplit.Length > 6)
            {
                string[] DataSourceName = connSplit[0].Split('=');
                string[] DataBaseName = connSplit[1].Split('=');
                string[] UserName = connSplit[2].Split('=');
                string[] PassWord = connSplit[3].Split('=');
                Program.DataBaseName = DataBaseName[1].ToString();
                Program.ServerName = DataSourceName[1].ToString();
                Program.ServerUserId = UserName[1].ToString();
                Program.ServerPassword = PassWord[1].ToString();
            }

            Application.Run(new form_TallyExport());
        }
    }
}
