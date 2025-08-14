using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace pcn_box
{
    static class Program
    {
        public static string connstring;
        public static SqlConnection conn;
        public static List<string> list_obj_type;// = new List<string>();
        public static List<string> list_protect_type;// = new List<string>();
        public static int comments;

        public static string fast_str;

        public static bool sql_insert_deblo = false;

        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {

         
            int ip9_1 = 192;
            int ip9_2 = 168;
            int ip9_3 = 0;
            int ip9_4 = 161;


            if (File.Exists("conn.txt") && File.Exists("cm.txt"))
            {
                int pass_ext = 123;
                connstring = File.ReadAllText("conn.txt");
                conn = new SqlConnection("Data Source=" + ip9_1.ToString() + "." + ip9_2.ToString() + "." + ip9_3.ToString() + "." +ip9_4.ToString() + ";" + connstring + "nXt" + pass_ext.ToString() + pass_ext.ToString());

                try
                {
                    conn.Open();
                }
                catch (SqlException sql_ex)
                {
                    StringBuilder errorMessages = new StringBuilder();

                    for (int i = 0; i < sql_ex.Errors.Count; i++)
                    {
                        errorMessages.Append("Index #" + i + "\n" +
                            "Message: " + sql_ex.Errors[i].Message + "\n" +
                            "Error Number: " + sql_ex.Errors[i].Number + "\n" +
                            "LineNumber: " + sql_ex.Errors[i].LineNumber + "\n" +
                            "Source: " + sql_ex.Errors[i].Source + "\n" +
                            "Procedure: " + sql_ex.Errors[i].Procedure + "\n");
                    }
                    MessageBox.Show("Проблемы связи с сервером: SQL Server - " + errorMessages.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Проблемы связи с сервером: SQL Server - class:" + ex.Message);
                }

            }
            else {
                if (File.Exists("conn.txt"))
                {
                    int pass_ext = 123;
                    connstring = File.ReadAllText("conn.txt");
                    //conn = new SqlConnection(connstring + "nXt" + pass_ext.ToString() + pass_ext.ToString());
                    conn = new SqlConnection("Data Source=" + ip1_1.ToString() + "." + ip1_2.ToString() + "." + ip1_3.ToString() + "." + ip1_4.ToString() + ";" + connstring + "nXt" + pass_ext.ToString() + pass_ext.ToString());
                }

                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {

                    int pass_ext = 123;
                    connstring = File.ReadAllText("conn.txt");
                    //conn = new SqlConnection(connstring + "nXt" + pass_ext.ToString() + pass_ext.ToString());
                    conn = new SqlConnection("Data Source=" + ip2_1.ToString() + "." + ip2_2.ToString() + "." + ip2_3.ToString() + "." + ip2_4.ToString() + ";" + connstring + "nXt" + pass_ext.ToString() + pass_ext.ToString());

                    try
                    {
                        conn.Open();
                    }
                    catch (SqlException sql_ex)
                    {


                        StringBuilder errorMessages = new StringBuilder();

                        for (int i = 0; i < sql_ex.Errors.Count; i++)
                        {
                            errorMessages.Append("Index #" + i + "\n" +
                                "Message: " + sql_ex.Errors[i].Message + "\n" +
                                "Error Number: " + sql_ex.Errors[i].Number + "\n" +
                                "LineNumber: " + sql_ex.Errors[i].LineNumber + "\n" +
                                "Source: " + sql_ex.Errors[i].Source + "\n" +
                                "Procedure: " + sql_ex.Errors[i].Procedure + "\n");
                        }
                        MessageBox.Show("Проблемы связи с сервером: SQL Server - " + errorMessages.ToString());
                    }
                    catch (Exception ex1)
                    {
                        MessageBox.Show("Проблемы связи с сервером: SQL Server - class:" + ex1.Message);
                    }
                }
            }

            list_obj_type = new List<string>();
            list_protect_type = new List<string>();
            fast_str="";

            comments=0;



            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
