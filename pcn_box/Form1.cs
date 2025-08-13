using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Diagnostics;
//using "";

namespace pcn_box
{
    public partial class Form1 : Form
    {
        //Thread RefreshDB_Disl;
        private string imp_err;
        private int rep_time;
        public static string xls_path;

        public Form1()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        private void обновитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ThreadStart thr1 = new ThreadStart(Refresh_DB_Disl);
            Thread RefreshDB_Disl = new Thread(thr1);
            RefreshDB_Disl.Start();


        }
        public void Refresh_DB_Disl()
        {
            List<string> ini_disl_obj = new List<string>();
            object[,] arr1 = new object[5001, 1];
            object[,] arr2 = new object[5001, 1];
            object[,] arr3 = new object[5001, 1];
            object[,] arr4 = new object[5001, 1];
            object[,] arr5 = new object[5001, 1];
            object[,] arr6 = new object[5001, 1];
            object[,] arr7 = new object[5001, 1];
            object[,] arr8 = new object[5001, 1];
            object[,] arr9 = new object[5001, 1];
            object[,] arr10 = new object[5001, 1];
            object[,] arr11 = new object[5001, 1];

            List<string> ini_disl_lig = new List<string>();
            object[,] arrl1 = new object[5001, 1];
            object[,] arrl2 = new object[5001, 1];
            object[,] arrl3 = new object[5001, 1];
            object[,] arrl4 = new object[5001, 1];
            object[,] arrl5 = new object[5001, 1];
            object[,] arrl6 = new object[5001, 1];
            object[,] arrl7 = new object[5001, 1];
            object[,] arrl8 = new object[5001, 1];
            object[,] arrl9 = new object[5001, 1];
            object[,] arrl10 = new object[5001, 1];


            if (File.Exists("path_disl_obj.txt"))
            {
                ini_disl_obj.AddRange((string[])File.ReadAllLines("path_disl_obj.txt", Encoding.GetEncoding(1251)));
            }
            toolStripProgressBar1.Value = 5;


            if (File.Exists("path_disl_lig.txt"))
            {
                ini_disl_lig.AddRange((string[])File.ReadAllLines("path_disl_lig.txt", Encoding.GetEncoding(1251)));
            }

            if (!File.Exists(ini_disl_obj[0]) || !File.Exists(ini_disl_lig[0]))
            {
                MessageBox.Show("Нет файла: %дислокация объектов% \n Нет файла: %дислокация ЛИГов% ");
            }
            else {

                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(ini_disl_obj[0], 0, true, 5, "907", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                toolStripProgressBar1.Value = 10;

                groupBox10.Visible = true;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = 5000;

                Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range(ini_disl_obj[1] + "1", ini_disl_obj[1] + "5000");
                arr1 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[2] + "1", ini_disl_obj[2] + "5000");
                arr2 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[3] + "1", ini_disl_obj[3] + "5000");
                arr3 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[4] + "1", ini_disl_obj[4] + "5000");
                arr4 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[5] + "1", ini_disl_obj[5] + "5000");
                arr5 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[6] + "1", ini_disl_obj[6] + "5000");
                arr6 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[7] + "1", ini_disl_obj[7] + "5000");
                arr7 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[8] + "1", ini_disl_obj[8] + "5000");
                arr8 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[9] + "1", ini_disl_obj[9] + "5000");
                arr9 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[10] + "1", ini_disl_obj[10] + "5000");
                arr10 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_obj[11] + "1", ini_disl_obj[11] + "5000");
                arr11 = (object[,])range.Value2;

                ///////////////////////////////////////////////////ObjWorkBook.Close(null, null, null);

                ObjWorkBook.Close(null, null, null);

                Console.WriteLine("xl_read_end");

                toolStripProgressBar1.Value = 15;

                SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                string sql_comm_str;

                sql_comm_str = "TRUNCATE TABLE [dbo].[uDisl]";

                comm2.CommandText = sql_comm_str;
                comm2.ExecuteNonQuery();

                progressBar1.Minimum = 0;
                progressBar1.Maximum = 5000;


                for (int i = 1; i < 5000; i++)
                {
                    progressBar1.Value++;
                    double progress1;
                    if (arr2[i, 1] != null)
                    {
                        int type_protect = 0;
                        if (arr5[i, 1] != null)
                        {
                            if (arr5[i, 1].ToString().Contains("ТС"))
                                type_protect = 2;
                            else
                                type_protect = 1;
                        }
                        else
                            type_protect = 1;

                        if (arr1[i, 1] == null) arr1[i, 1] = "(пустое поле)";
                        if (arr3[i, 1] == null) arr3[i, 1] = "(пустое поле)";
                        if (arr4[i, 1] == null) arr4[i, 1] = "(пустое поле)";
                        if (arr6[i, 1] == null) arr6[i, 1] = "(пусто)";
                        if (arr7[i, 1] == null) arr7[i, 1] = "(пусто)";
                        if (arr8[i, 1] == null) arr8[i, 1] = "-";
                        if (arr9[i, 1] == null) arr9[i, 1] = "-";
                        if (arr10[i, 1] == null) arr10[i, 1] = "нет";
                        if (arr11[i, 1] == null) arr11[i, 1] = "0,0";


                        string temp_str2 = arr7[i, 1].ToString().Trim();
                        if (temp_str2.Contains(' '))
                        {
                            string str_wo_ws = "";
                            for (int s = 0; s < temp_str2.Length; s++)
                            {
                                if (temp_str2[s] == '(')
                                    break;
                                if (temp_str2[s] != ' ')
                                    str_wo_ws += temp_str2[s];
                            }
                            arr7[i, 1] = str_wo_ws;

                        }

                        if (arr1[i, 1].ToString().Contains("'")) arr1[i, 1] = arr1[i, 1].ToString().Replace('\'', ' ');
                        if (arr2[i, 1].ToString().Contains("'")) arr2[i, 1] = arr2[i, 1].ToString().Replace('\'', ' ');
                        if (arr3[i, 1].ToString().Contains("'")) arr3[i, 1] = arr3[i, 1].ToString().Replace('\'', ' ');
                        if (arr4[i, 1].ToString().Contains("'")) arr4[i, 1] = arr4[i, 1].ToString().Replace('\'', ' ');
                        if (arr6[i, 1].ToString().Contains("'")) arr6[i, 1] = arr6[i, 1].ToString().Replace('\'', ' ');
                        if (arr7[i, 1].ToString().Contains("'")) arr7[i, 1] = arr7[i, 1].ToString().Replace('\'', ' ');
                        if (arr8[i, 1].ToString().Contains("'")) arr8[i, 1] = arr8[i, 1].ToString().Replace('\'', ' ');
                        if (arr9[i, 1].ToString().Contains("'")) arr9[i, 1] = arr9[i, 1].ToString().Replace('\'', ' ');
                        if (arr10[i, 1].ToString().Contains("'")) arr10[i, 1] = arr10[i, 1].ToString().Replace('\'', ' ');
                        if (arr11[i, 1].ToString().Contains("'")) arr11[i, 1] = arr11[i, 1].ToString().Replace('\'', ' ');


                        sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[serv_org], [dogovor], [tel_info], [deblo_otvetish], [average_month]) VALUES ('" + arr1[i, 1].ToString() + "', '" + arr2[i, 1].ToString() + "', '" + arr3[i, 1].ToString() + "','" + arr4[i, 1].ToString() + "',1," + type_protect + ", '" + arr6[i, 1].ToString() + "', '" + arr7[i, 1].ToString() + "', '" + arr8[i, 1].ToString() + ", " + arr9[i, 1].ToString() + "', '" + arr10[i, 1].ToString() + "', '" + arr11[i, 1].ToString() + "')";
                        //SqlCommand comm2 = new SqlCommand(sql_comm_str, Program.conn);
                        comm2.CommandText = sql_comm_str;
                        comm2.ExecuteNonQuery();
                        //Dispose()
                    }
                    progress1 = i / 5000;
                    toolStripProgressBar1.Value = 15 + (int)(20 * progress1);
                }
                toolStripProgressBar1.Value = 49;
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                List<string> ini_disl_obj_pak = new List<string>();
                object[,] arrp1 = new object[5001, 1];
                object[,] arrp2 = new object[5001, 1];
                object[,] arrp3 = new object[5001, 1];
                object[,] arrp4 = new object[5001, 1];
                object[,] arrp5 = new object[5001, 1];
                object[,] arrp6 = new object[5001, 1];

                if (File.Exists("path_disl_obj_pak.txt"))
                {
                    ini_disl_obj_pak.AddRange((string[])File.ReadAllLines("path_disl_obj_pak.txt", Encoding.GetEncoding(1251)));
                }
                toolStripProgressBar1.Value = 5;
                /*
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                */
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookPAK = ObjExcel.Workbooks.Open(ini_disl_obj_pak[0], 0, true, 5, "907", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet_pak;
                ObjWorkSheet_pak = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookPAK.Sheets[1];

                toolStripProgressBar1.Value = 10;

                Microsoft.Office.Interop.Excel.Range rangeP = ObjWorkSheet_pak.get_Range(ini_disl_obj_pak[1] + "1", ini_disl_obj_pak[1] + "5000");
                arrp1 = (object[,])rangeP.Value2;
                rangeP = ObjWorkSheet_pak.get_Range(ini_disl_obj_pak[2] + "1", ini_disl_obj_pak[2] + "5000");
                arrp2 = (object[,])rangeP.Value2;
                rangeP = ObjWorkSheet_pak.get_Range(ini_disl_obj_pak[3] + "1", ini_disl_obj_pak[3] + "5000");
                arrp3 = (object[,])rangeP.Value2;
                rangeP = ObjWorkSheet_pak.get_Range(ini_disl_obj_pak[4] + "1", ini_disl_obj_pak[4] + "5000");
                arrp4 = (object[,])rangeP.Value2;
                rangeP = ObjWorkSheet_pak.get_Range(ini_disl_obj_pak[5] + "1", ini_disl_obj_pak[5] + "5000");
                arrp5 = (object[,])rangeP.Value2;
                //rangeP = ObjWorkSheet_pak.get_Range(ini_disl_obj_pak[6] + "1", ini_disl_obj_pak[6] + "5000");
                //arrp6 = (object[,])rangeP.Value2;

                ObjWorkBookPAK.Close(null, null, null);

                Console.WriteLine("xl_read_end");

                toolStripProgressBar1.Value = 15;

                /*SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                string sql_comm_str;

                sql_comm_str = "TRUNCATE TABLE [dbo].[uDisl]";

                comm2.CommandText = sql_comm_str;
                comm2.ExecuteNonQuery();
                */

                progressBar2.Minimum = 0;
                progressBar2.Maximum = 5000;

                for (int i = 1; i < 5000; i++)
                {
                    progressBar2.Value++;
                    double progress1;
                    if (arrp2[i, 1] != null)
                    {
                        int type_protect = 0;
                        if (arrp5[i, 1] != null)
                        {
                            if (arrp5[i, 1].ToString().Contains("3"))
                                type_protect = 2;
                            else
                                type_protect = 1;
                        }
                        else
                            type_protect = 1;

                        if (arrp1[i, 1] == null) arrp1[i, 1] = "(пустое поле)";
                        if (arrp3[i, 1] == null) arrp3[i, 1] = "(пустое поле)";
                        if (arrp4[i, 1] == null) arrp4[i, 1] = "(пустое поле)";
                        //if (arrp6[i, 1] == null) arrp6[i, 1] = "(пусто)";

                        if (arrp1[i, 1].ToString().Contains("'")) arrp1[i, 1] = arrp1[i, 1].ToString().Replace('\'', ' ');
                        if (arrp2[i, 1].ToString().Contains("'")) arrp2[i, 1] = arrp2[i, 1].ToString().Replace('\'', ' ');
                        if (arrp3[i, 1].ToString().Contains("'")) arrp3[i, 1] = arrp3[i, 1].ToString().Replace('\'', ' ');
                        if (arrp4[i, 1].ToString().Contains("'")) arrp4[i, 1] = arrp4[i, 1].ToString().Replace('\'', ' ');
                        //if (arrp6[i, 1].ToString().Contains("'")) arrp6[i, 1] = arrp6[i, 1].ToString().Replace('\'', ' ');


                        sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[serv_org]) VALUES ('" + arrp1[i, 1].ToString() + "', '" + arrp2[i, 1].ToString() + "', '" + arrp3[i, 1].ToString() + "','" + arrp4[i, 1].ToString() + "',4," + type_protect + ", 'ФГУП')";

                        comm2.CommandText = sql_comm_str;
                        comm2.ExecuteNonQuery();

                    }
                    progress1 = i / 5000;
                    toolStripProgressBar1.Value = 15 + (int)(20 * progress1);
                }
                toolStripProgressBar1.Value = 49;

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                toolStripProgressBar1.Value = 50;


                //Создаём приложение.
                //Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookL = ObjExcel.Workbooks.Open(ini_disl_lig[0], 0, true, 5, "907", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetL;
                ObjWorkSheetL = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookL.Sheets[1];

                toolStripProgressBar1.Value = 60;

                progressBar1.Minimum = 0;
                progressBar1.Maximum = 5000;
                progressBar1.Value = 0;


                Microsoft.Office.Interop.Excel.Range rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[1] + "1", ini_disl_lig[1] + "5000");
                arrl1 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[2] + "1", ini_disl_lig[2] + "5000");
                arrl2 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[3] + "1", ini_disl_lig[3] + "5000");
                arrl3 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[4] + "1", ini_disl_lig[4] + "5000");
                arrl4 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[5] + "1", ini_disl_lig[5] + "5000");
                arrl5 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[6] + "1", ini_disl_lig[6] + "5000");
                arrl6 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[7] + "1", ini_disl_lig[7] + "5000");
                arrl7 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[8] + "1", ini_disl_lig[8] + "5000");
                arrl8 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[9] + "1", ini_disl_lig[9] + "5000");
                arrl9 = (object[,])rangeL.Value2;
                rangeL = ObjWorkSheetL.get_Range(ini_disl_lig[10] + "1", ini_disl_lig[10] + "5000");
                arrl10 = (object[,])rangeL.Value2;

                ObjWorkBookL.Close(null, null, null);

                //////////////////////////////////////////////////////////////ObjWorkBookL.Close(null, null, null);

                Console.WriteLine("xl_read_end");

                toolStripProgressBar1.Value = 65;


                for (int i = 1; i < 5000; i++)
                {
                    progressBar1.Value++;
                    double progress1;
                    if (arrl2[i, 1] != null)
                    {
                        int type_protect = 0;
                        if (arrl5[i, 1] != null)
                        {
                            if (arrl5[i, 1].ToString().Contains('3'))
                                type_protect = 2;
                            else
                                type_protect = 1;
                        }
                        else
                            type_protect = 1;

                        if (arrl1[i, 1] == null) arrl1[i, 1] = "(пустое поле)";
                        if (arrl3[i, 1] == null) arrl3[i, 1] = "(пустое поле)";
                        if (arrl4[i, 1] == null) arrl4[i, 1] = "(пустое поле)";
                        if (arrl6[i, 1] == null) arrl6[i, 1] = "(пусто)";
                        if (arrl7[i, 1] == null) arrl7[i, 1] = "(пусто)";
                        if (arrl8[i, 1] == null) arrl8[i, 1] = "-";
                        if (arrl9[i, 1] == null) arrl9[i, 1] = "-";
                        if (arrl9[i, 1].ToString().Trim().Length > 6) arrl9[i, 1] = "-";
                        if (arrl10[i, 1] == null) arrl10[i, 1] = "нет";

                        string temp_str2 = arrl7[i, 1].ToString().Trim();
                        if (temp_str2.Contains(' '))
                        {
                            string str_wo_ws = "";
                            for (int s = 0; s < temp_str2.Length; s++)
                            {
                                if (temp_str2[s] == '(')
                                    break;
                                if (temp_str2[s] != ' ')
                                    str_wo_ws += temp_str2[s];
                            }
                            arrl7[i, 1] = str_wo_ws;

                        }


                        if (arrl1[i, 1].ToString().Contains("'")) arrl1[i, 1] = arrl1[i, 1].ToString().Replace('\'', ' ');
                        if (arrl2[i, 1].ToString().Contains("'")) arrl2[i, 1] = arrl2[i, 1].ToString().Replace('\'', ' ');
                        if (arrl3[i, 1].ToString().Contains("'")) arrl3[i, 1] = arrl3[i, 1].ToString().Replace('\'', ' ');
                        if (arrl4[i, 1].ToString().Contains("'")) arrl4[i, 1] = arrl4[i, 1].ToString().Replace('\'', ' ');
                        if (arrl6[i, 1].ToString().Contains("'")) arrl6[i, 1] = arrl6[i, 1].ToString().Replace('\'', ' ');
                        if (arrl7[i, 1].ToString().Contains("'")) arrl7[i, 1] = arrl7[i, 1].ToString().Replace('\'', ' ');
                        if (arrl8[i, 1].ToString().Contains("'")) arrl8[i, 1] = arrl8[i, 1].ToString().Replace('\'', ' ');
                        if (arrl9[i, 1].ToString().Contains("'")) arrl9[i, 1] = arrl9[i, 1].ToString().Replace('\'', ' ');
                        if (arrl10[i, 1].ToString().Contains("'")) arrl10[i, 1] = arrl10[i, 1].ToString().Replace('\'', ' ');


                        int type_LIG = 0;
                        if (arrl3[i, 1] != null)
                        {
                            if (arrl3[i, 1].ToString().ToLower() == "квартира")
                            {
                                type_LIG = 2;
                            }
                            else
                            {
                                type_LIG = 3;
                            }
                        }
                        else
                            type_LIG = 2;


                        //sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[serv_org]) VALUES ('" + arrl1[i, 1].ToString() + "', '" + arrl2[i, 1].ToString() + "', '" + arrl3[i, 1].ToString() + "','" + arrl4[i, 1].ToString() + "'," + type_LIG + "," + type_protect + ", '" + arrl6[i, 1].ToString() + "')";

                        //sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[serv_org], [dogovor], [tel_info], [deblo_otvetish]) VALUES ('" + arrl1[i, 1].ToString() + "', '" + arrl2[i, 1].ToString() + "', '" + arrl3[i, 1].ToString() + "','" + arrl4[i, 1].ToString() + "'," + type_LIG + "," + type_protect + ", '" + arrl6[i, 1].ToString() + "', '" + arrl7[i, 1].ToString() + "', '" + arrl8[i, 1].ToString() + "', '" + arrl9[i, 1].ToString() + "', '" + arrl10[i, 1].ToString() + "')";
                        sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[serv_org], [dogovor], [tel_info], [deblo_otvetish]) VALUES ('" + arrl1[i, 1].ToString() + "', '" + arrl2[i, 1].ToString() + "', '" + arrl3[i, 1].ToString() + "','" + arrl4[i, 1].ToString() + "'," + type_LIG + "," + type_protect + ", '" + arrl6[i, 1].ToString() + "', '" + arrl7[i, 1].ToString() + "', '" + arrl8[i, 1].ToString() + ", " + arrl9[i, 1].ToString() + "', '" + arrl10[i, 1].ToString() + "')";
                        //SqlCommand comm2 = new SqlCommand(sql_comm_str, Program.conn);
                        comm2.CommandText = sql_comm_str;
                        comm2.ExecuteNonQuery();
                        //Dispose()
                    }
                    progress1 = i / 5000;
                    toolStripProgressBar1.Value = 65 + (int)(20 * progress1);
                }
                /////////////////////////////////////////////////////////////////////////////////////////////////////0
                List<string> ini_disl_lig_pak = new List<string>();
                object[,] arrlp1 = new object[5001, 1];
                object[,] arrlp2 = new object[5001, 1];
                object[,] arrlp3 = new object[5001, 1];
                object[,] arrlp4 = new object[5001, 1];
                object[,] arrlp5 = new object[5001, 1];
                object[,] arrlp6 = new object[5001, 1];

                if (File.Exists("path_disl_lig_pak.txt"))
                {
                    ini_disl_lig_pak.AddRange((string[])File.ReadAllLines("path_disl_lig_pak.txt", Encoding.GetEncoding(1251)));
                }
                toolStripProgressBar1.Value = 50;

                //Создаём приложение.
                //Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookLPAK = ObjExcel.Workbooks.Open(ini_disl_lig_pak[0], 0, true, 5, "907", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetLP;
                ObjWorkSheetLP = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookLPAK.Sheets[1];

                toolStripProgressBar1.Value = 60;

                progressBar2.Minimum = 0;
                progressBar2.Maximum = 5000;
                progressBar2.Value = 0;


                Microsoft.Office.Interop.Excel.Range rangeLP = ObjWorkSheetLP.get_Range(ini_disl_lig_pak[1] + "1", ini_disl_lig_pak[1] + "5000");
                arrlp1 = (object[,])rangeLP.Value2;
                rangeLP = ObjWorkSheetLP.get_Range(ini_disl_lig_pak[2] + "1", ini_disl_lig_pak[2] + "5000");
                arrlp2 = (object[,])rangeLP.Value2;
                rangeLP = ObjWorkSheetLP.get_Range(ini_disl_lig_pak[3] + "1", ini_disl_lig_pak[3] + "5000");
                arrlp3 = (object[,])rangeLP.Value2;
                rangeLP = ObjWorkSheetLP.get_Range(ini_disl_lig_pak[4] + "1", ini_disl_lig_pak[4] + "5000");
                arrlp4 = (object[,])rangeLP.Value2;
                rangeLP = ObjWorkSheetLP.get_Range(ini_disl_lig_pak[5] + "1", ini_disl_lig_pak[5] + "5000");
                arrlp5 = (object[,])rangeLP.Value2;
                //rangeLP = ObjWorkSheetLP.get_Range(ini_disl_lig_pak[6] + "1", ini_disl_lig_pak[6] + "5000");
                //arrlp6 = (object[,])rangeLP.Value2;

                ObjWorkBookLPAK.Close(null, null, null);

                Console.WriteLine("xl_read_end");

                toolStripProgressBar1.Value = 65;


                for (int i = 1; i < 5000; i++)
                {
                    progressBar2.Value++;
                    double progress1;
                    if (arrlp2[i, 1] != null)
                    {
                        int type_protect = 0;
                        if (arrlp5[i, 1] != null)
                        {
                            if (arrlp5[i, 1].ToString().Contains('3'))
                                type_protect = 2;
                            else
                                type_protect = 1;
                        }
                        else
                            type_protect = 1;

                        if (arrlp1[i, 1] == null) arrlp1[i, 1] = "(пустое поле)";
                        if (arrlp3[i, 1] == null) arrlp3[i, 1] = "(пустое поле)";
                        if (arrlp4[i, 1] == null) arrlp4[i, 1] = "(пустое поле)";
                        //if (arrlp6[i, 1] == null) arrlp6[i, 1] = "(пусто)";

                        if (arrlp1[i, 1].ToString().Contains("'")) arrlp1[i, 1] = arrlp1[i, 1].ToString().Replace('\'', ' ');
                        if (arrlp2[i, 1].ToString().Contains("'")) arrlp2[i, 1] = arrlp2[i, 1].ToString().Replace('\'', ' ');
                        if (arrlp3[i, 1].ToString().Contains("'")) arrlp3[i, 1] = arrlp3[i, 1].ToString().Replace('\'', ' ');
                        if (arrlp4[i, 1].ToString().Contains("'")) arrlp4[i, 1] = arrlp4[i, 1].ToString().Replace('\'', ' ');
                        //if (arrlp6[i, 1].ToString().Contains("'")) arrlp6[i, 1] = arrlp6[i, 1].ToString().Replace('\'', ' ');

                        /*int type_LIG = 0;
                        if (arrlp3[i, 1] != null)
                        {
                            if (arrlp3[i, 1].ToString().ToLower() == "квартира")
                            {
                                type_LIG = 2;
                            }
                            else
                            {
                                type_LIG = 3;
                            }
                        }
                        else
                            type_LIG = 2;*/

                        sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[serv_org]) VALUES ('" + arrlp1[i, 1].ToString() + "', '" + arrlp2[i, 1].ToString() + "', '" + arrlp3[i, 1].ToString() + "','" + arrlp4[i, 1].ToString() + "'," + 5 + "," + type_protect + ", 'ФГУП')";
                        //SqlCommand comm2 = new SqlCommand(sql_comm_str, Program.conn);
                        comm2.CommandText = sql_comm_str;
                        comm2.ExecuteNonQuery();
                        //Dispose()
                    }
                    progress1 = i / 5000;
                    toolStripProgressBar1.Value = 65 + (int)(20 * progress1);
                }

                toolStripProgressBar1.Value = 95;
                this.uDislTableAdapter.Fill(this.uniq1DataSet.uDisl);
                progressBar1.Value = 0; progressBar2.Value = 0;
                groupBox10.Visible = false;
            }
            toolStripProgressBar1.Value = 100;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel1.Width = 268;
            panel1.Height = 1242;

            //Data Source=PCN921\sqlserver12;Initial Catalog=uniq1;Persist Security Info=True;User ID=uu;Password=nXt321321

            // TODO: данная строка кода позволяет загрузить данные в таблицу "uniq1DataSet.uDisl". При необходимости она может быть перемещена или удалена.
            this.uDislTableAdapter.Connection = Program.conn;
            this.uDislTableAdapter.Fill(this.uniq1DataSet.uDisl);
            label1.Text = "Количество объектов: " + dataGridView1.RowCount.ToString();
            // TODO: данная строка кода позволяет загрузить данные в таблицу "uniq1DataSet.uDep". При необходимости она может быть перемещена или удалена.
            this.uDepTableAdapter.Connection = Program.conn;
            //SELECT TOP (3) idd, poz, obj_org, obj_name, obj_adr, obj_type, type_protect, date_report, date_alarm, date_pass, date_arrive, date_complete, gz, gz_start, comment_alarm, comment_work, comment_afterwork, comment_general, reason_start, reason_end, worker1, worker2, worker3, worker4, service_org_name, import_errors, e_ment, temper, weather, date_afterwork, dogovor1, comment_edit, average_month, last_month FROM     uDep ORDER BY idd DESC

            if (File.Exists("up_sql1.txt"))
            {
                //
                string sql_comm_str = "";
                string xXx = "";
                string yYy = "";

                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up1")
                {
                    //10. УТЗ
                    xXx = "10. УТЗ"; // new
                    yYy = "9.УТЗ";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up2")
                {
                    //9. Другие
                    xXx = "9. Другие"; // new
                    yYy = "6. Другие"; // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up3")
                {
                    //8.5 Вина. Другие
                    xXx = "8.5 Вина. Другие"; // new
                    yYy = "4.5. (Вина) Другие";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up4")
                {
                    //8.4 Вина. Животные
                    xXx = "8.4 Вина. Животные"; // new
                    yYy = "4.4.(Вина)Животные, насек";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up5")
                {
                    //8.3 Вина. Случайно ТС
                    xXx = "8.3 Вина. Случайно ТС"; // new
                    yYy = "4.3. (Вина) Случайно ТС";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up6")
                {
                    //8.2 Вина. Неверн. действ.
                    xXx = "8.2 Вина. Неверн. действ."; // new
                    yYy = "4.2. (Вина) Неверн. дейст";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up7")
                {
                    //8.1 Вина. Забыли снять
                    xXx = "8.1 Вина. Забыли снять"; // new
                    yYy = "4.1. (Вина) Забыли снять";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up8")
                {
                    //7. Неудовлетв. ИТУ
                    xXx = "7. Неудовлетв. ИТУ"; // new
                    yYy = "Тех. укрепленность";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up9")
                {
                    //6. Обестачивание
                    xXx = "6. Обестачивание"; // new
                    yYy = "3.Обестачивание";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up10")
                {
                    //5. Вина АТС************************************************************
                    xXx = "5. Вина АТС"; // new
                    yYy = "2.Вина АТС";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up11")
                {
                    //5. Вина АТС************************************************************
                    xXx = "5. Вина АТС"; // new
                    yYy = "2.Вина оператора связи";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up12")
                {
                    //3. Неиспр. ППК
                    xXx = "3. Неиспр. ППК"; // new
                    yYy = "Неиспр. ППК";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up13")
                {
                    //2. Неиспр. датчики
                    xXx = "2. Неиспр. датчики"; // new
                    yYy = "Неиспр. Датчики";   // old
                }
                if (File.ReadAllLines("up_sql1.txt", Encoding.GetEncoding(1251))[0] == "up14")
                {
                    //1. Неиспр. МКИ и ШС
                    xXx = "1. Неиспр. МКИ и ШС"; // new
                    yYy = "Неиспр. Шлейфы, СМК";   // old
                }

                SqlCommand imp_dep = new SqlCommand();
                imp_dep.Connection = Program.conn;
                sql_comm_str = "UPDATE [dbo].[uDep] SET [reason_start]='" + xXx + "' WHERE reason_start='" + yYy + "' ";
                imp_dep.CommandText = sql_comm_str;
                imp_dep.ExecuteNonQuery();

                SqlCommand imp_dep2 = new SqlCommand();
                imp_dep2.Connection = Program.conn;
                sql_comm_str = "UPDATE [dbo].[uDep] SET [reason_end]='" + xXx + "' WHERE reason_end='" + yYy + "' ";
                imp_dep2.CommandText = sql_comm_str;
                imp_dep2.ExecuteNonQuery();

            }
            if (File.Exists("up_sql2.txt"))
            {
                //
                string sql_comm_str = "";
                if (File.ReadAllLines("up_sql2.txt", Encoding.GetEncoding(1251))[0] == "up1")
                {
                    SqlCommand comm2 = new SqlCommand();
                    comm2.Connection = Program.conn;

                    sql_comm_str = @"SELECT id,poz FROM  [dbo].[uJupiter8] WHERE [id]>3357 ORDER BY [id]";//3357 !!! great thanks to kolobok !!!
                    comm2.CommandText = sql_comm_str;
                    SqlDataReader test_r1 = comm2.ExecuteReader();

                    object[,] ds1 = new object[15001, 2];
                    int rowcount = 0;
                    while (test_r1.Read())
                    {
                        ds1[rowcount, 0] = test_r1[0];
                        ds1[rowcount, 1] = test_r1[1];

                        if (ds1[rowcount, 1].ToString().Length == 1)
                            ds1[rowcount, 1] = (object)("000" + ds1[rowcount, 1].ToString());
                        if (ds1[rowcount, 1].ToString().Length == 2)
                            ds1[rowcount, 1] = (object)("00" + ds1[rowcount, 1].ToString());
                        if (ds1[rowcount, 1].ToString().Length == 3)
                            ds1[rowcount, 1] = (object)("0" + ds1[rowcount, 1].ToString());

                        rowcount++;
                    }
                    test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////
                    int i = 0;
                    for (i = 0; i < rowcount; i++)
                    {
                        SqlCommand imp_dep = new SqlCommand();
                        imp_dep.Connection = Program.conn;
                        sql_comm_str = "UPDATE [dbo].[uJupiter8] SET [poz]='" + ds1[i, 1].ToString() + "' WHERE id='" + ds1[i, 0].ToString() + "' ";
                        imp_dep.CommandText = sql_comm_str;
                        imp_dep.ExecuteNonQuery();

                    }
                    MessageBox.Show("up1_i="+i.ToString());
                }
                if (File.ReadAllLines("up_sql2.txt", Encoding.GetEncoding(1251))[0] == "up2")
                {
                    SqlCommand comm2 = new SqlCommand();
                    comm2.Connection = Program.conn;

                    sql_comm_str = @"SELECT id,poz FROM  [dbo].[uJupiterGOLD] WHERE [id]>0 ORDER BY [id]";
                    comm2.CommandText = sql_comm_str;
                    SqlDataReader test_r1 = comm2.ExecuteReader();

                    object[,] ds1 = new object[15001, 2];
                    int rowcount = 0;
                    while (test_r1.Read())
                    {
                        ds1[rowcount, 0] = test_r1[0];
                        ds1[rowcount, 1] = test_r1[1];

                        if (ds1[rowcount, 1].ToString().Length == 1)
                            ds1[rowcount, 1] = (object)("000" + ds1[rowcount, 1].ToString());
                        if (ds1[rowcount, 1].ToString().Length == 2)
                            ds1[rowcount, 1] = (object)("00" + ds1[rowcount, 1].ToString());
                        if (ds1[rowcount, 1].ToString().Length == 3)
                            ds1[rowcount, 1] = (object)("0" + ds1[rowcount, 1].ToString());

                        rowcount++;
                    }
                    test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////
                    int i = 0;
                    for (i = 0; i < rowcount; i++)
                    {
                        SqlCommand imp_dep = new SqlCommand();
                        imp_dep.Connection = Program.conn;
                        sql_comm_str = "UPDATE [dbo].[uJupiterGOLD] SET [poz]='" + ds1[i, 1].ToString() + "' WHERE id='" + ds1[i, 0].ToString() + "' ";
                        imp_dep.CommandText = sql_comm_str;
                        imp_dep.ExecuteNonQuery();

                    }
                    MessageBox.Show("up1_i=" + i.ToString());
                }

            }



            if (File.Exists("udep_row_count.txt"))
            {
                //
                this.uDepTableAdapter.FillBy(this.uniq1DataSet.uDep, Int32.Parse(File.ReadAllLines("udep_row_count.txt", Encoding.GetEncoding(1251))[0]));///!!!
            }
            else
                this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);

            label2.Text = "Количество сработок: " + dataGridView2.RowCount.ToString();

            //dataGridView2.BackgroundColor = Color.LightGray;
            dataGridView2.DefaultCellStyle.BackColor = Color.LightGray;

            label38.Visible = false;
            dateTimePicker10.Visible = false;


            if (File.Exists("obj_type.txt"))
            {
                comboBox1.Items.AddRange((object[])File.ReadAllLines("obj_type.txt", Encoding.GetEncoding(1251)));
                Program.list_obj_type.AddRange((string[])File.ReadAllLines("obj_type.txt", Encoding.GetEncoding(1251)));

            }
            if (File.Exists("protect_type.txt"))
            {
                comboBox2.Items.AddRange((object[])File.ReadAllLines("protect_type.txt", Encoding.GetEncoding(1251)));
                Program.list_protect_type.AddRange((string[])File.ReadAllLines("protect_type.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("gz.txt"))
            {
                comboBox5.Items.AddRange((object[])File.ReadAllLines("gz.txt", Encoding.GetEncoding(1251)));
                //Program.list_protect_type.AddRange((string[])File.ReadAllLines("protect_type.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("reason.txt"))
            {
                //comboBox3.Items.AddRange((object[])File.ReadAllLines("reason.txt", Encoding.GetEncoding(1251)));
                //comboBox4.Items.AddRange((object[])File.ReadAllLines("reason.txt", Encoding.GetEncoding(1251)));
                string TFS_tech = "4ever";

            }
            if (File.Exists("workers_d.txt"))
            {
                comboBox7.Items.AddRange((object[])File.ReadAllLines("workers_d.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("workers_e.txt"))
            {
                comboBox9.Items.AddRange((object[])File.ReadAllLines("workers_e.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("workers_s.txt"))
            {
                comboBox8.Items.AddRange((object[])File.ReadAllLines("workers_s.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("workers_o.txt"))
            {
                comboBox10.Items.AddRange((object[])File.ReadAllLines("workers_o.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("serv_comp.txt"))
            {
                comboBox6.Items.AddRange((object[])File.ReadAllLines("serv_comp.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("ver.txt"))
            {
                groupBox6.Text = toolStripStatusLabel1.Text = "Версия: " + File.ReadAllLines("ver.txt", Encoding.GetEncoding(1251))[0];
            }

            if (File.Exists("report_time1.txt"))
            {
                try
                {
                    rep_time = Int32.Parse(File.ReadAllLines("report_time1.txt", Encoding.GetEncoding(1251))[0]);
                }
                catch
                {
                    rep_time = 6;
                }
            }
            if (File.Exists("weather.txt"))
            {
                comboBox11.Items.AddRange((object[])File.ReadAllLines("weather.txt", Encoding.GetEncoding(1251)));
            }
            if (File.Exists("xls_path.txt"))
            {
                xls_path = File.ReadAllText("xls_path.txt", Encoding.GetEncoding(1251));

            }



            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            checkBox9.Checked = false;
            groupBox10.Visible = false;
            дислокацияToolStripMenuItem.Visible = false;
            дебиторскаяToolStripMenuItem.Visible = false;
            дислокацияToolStripMenuItem1.Checked = false;
            дебиторскаяToolStripMenuItem1.Checked = false;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //
            if (checkBox10.Checked) textBox2.Text = textBox1.Text;

            try
            {
                string filt1 = "";
                if (checkBox1.Checked) filt1 += "poz LIKE '%" + textBox1.Text + "%' OR ";
                if (checkBox2.Checked) filt1 += "obj_org LIKE '%" + textBox1.Text + "%' OR ";
                if (checkBox3.Checked) filt1 += "obj_name LIKE '%" + textBox1.Text + "%' OR ";
                if (checkBox4.Checked) filt1 += "obj_adr LIKE '%" + textBox1.Text + "%' OR ";

                //uDislBindingSource.Filter = "obj_org LIKE '%"+textBox1.Text+"%'";

                //string tmpst1= filt1.PadRight(4);
                if (filt1.EndsWith("OR ")) filt1 = filt1.Remove(filt1.Length - 3);
                uDislBindingSource.Filter = filt1;

                label1.Text = "Количество объектов: " + dataGridView1.RowCount.ToString();
                toolStripStatusLabel2.Text = "Статус: ";
            }
            catch
            {
                toolStripStatusLabel2.Text = "Статус: Ошибка поиска...";
                textBox1.Text = textBox1.Text.Replace('*', ' ');
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            textBox1_TextChanged(null, null);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox1_TextChanged(null, null);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            textBox1_TextChanged(null, null);
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            textBox1_TextChanged(null, null);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string filt1 = "";
                if (checkBox8.Checked) filt1 += "poz LIKE '%" + textBox2.Text + "%' OR ";
                if (checkBox7.Checked) filt1 += "obj_org LIKE '%" + textBox2.Text + "%' OR ";
                if (checkBox6.Checked) filt1 += "obj_name LIKE '%" + textBox2.Text + "%' OR ";
                if (checkBox5.Checked) filt1 += "obj_adr LIKE '%" + textBox2.Text + "%' OR ";

                if (filt1.EndsWith("OR ")) filt1 = filt1.Remove(filt1.Length - 3);
                uDepBindingSource.Filter = filt1;

                label2.Text = "Количество объектов: " + dataGridView2.RowCount.ToString();
                toolStripStatusLabel2.Text = "Статус: ";
            }
            catch
            {
                toolStripStatusLabel2.Text = "Статус: Ошибка поиска...";
                textBox2.Text = textBox2.Text.Replace('*', ' ');
            }
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            //

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex >= 0 && radioButton1.Checked)
            {
                imp_err = "";
                imp_err += label32.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                imp_err += "^";
                imp_err += textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                imp_err += "^";
                imp_err += textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                imp_err += "^";
                imp_err += textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                imp_err += "^";
                imp_err += textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                imp_err += "^";
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "1")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[0]);///*****************
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "2")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[1]);///*****************
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "3")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[2]);///*****************
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "4")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[3]);///*****************
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "5")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[4]);///*****************
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "6")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[5]);///*****************


                imp_err += dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
                imp_err += "^";
                if (dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString() == "1")
                    comboBox2.SelectedIndex = comboBox2.FindString(Program.list_protect_type[0]);///*****************
                if (dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString() == "2")
                    comboBox2.SelectedIndex = comboBox2.FindString(Program.list_protect_type[1]);///*****************
                imp_err += dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
                imp_err += "";
                label34.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
                textBox15.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
                textBox16.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
                textBox17.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
                textBox16.Enabled = false; textBox17.Enabled = false;
                textBox20.Text = dataGridView1.Rows[e.RowIndex].Cells[12].Value.ToString();
                if (dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString() == "1" && dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString() == "1")
                    groupBox11.Enabled = true;
                else
                    groupBox11.Enabled = false;
            }


        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void импортИзБДDepletionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string olestring = "";
            string _old_db_path;



            if (File.Exists("old_db_path.txt"))
            {
                _old_db_path = File.ReadAllText("old_db_path.txt", Encoding.GetEncoding(1251));
                olestring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _old_db_path + ";Jet OLEDB:Database Password=Chf,fnsdfybz;";
            }

            OleDbConnection oleconn = new OleDbConnection(olestring);
            oleconn.Open();
            string sql_comm_str;

            OleDbCommand test1 = new OleDbCommand(@"SELECT     Depletions.[Number], Depletions.CallLetter, Depletions.[Object], Organizations.OrganizationName, Objects.ObjectName, Objects.Adress, 
                      ObjectTypes.ObjectType, ProtectionTypes.ProtectionType, Depletions.ReportDate, Depletions.DateTimeAlarm, Depletions.DateTimePass, 
                      OGZNumber.OGZNumber, Depletions.OGZStartPoint, Depletions.DateTimeArrive, Depletions.Description, Depletions.OnDutyWorks, 
                      Depletions.CompleteWorks, Depletions.InitialReason, Reasons.Reason AS Reasons_Reason, Depletions.FinalReason, 
                      Reasons2.Reason AS Reasons2_Reason, Depletions.Engineer, Engineers.FamilyNP AS Engineers_FamilyNP, Depletions.Electrician, 
                      Electricians.FamilyNP AS Electricians_FamilyNP, Depletions.OnDuty, OnDutys.FamilyNP AS OnDutys_FamilyNP, Depletions.ShiftEngineer, 
                      ShiftEngineers.FamilyNP AS ShiftEngineers_FamilyNP, Depletions.Commentary
            FROM      (ShiftEngineers INNER JOIN
                      (Reasons INNER JOIN
                      (Reasons2 INNER JOIN
                      (ProtectionTypes INNER JOIN
                      (Organizations INNER JOIN
                      (OnDutys INNER JOIN
                      (OGZNumber INNER JOIN
                      (ObjectTypes INNER JOIN
                      (Objects INNER JOIN
                      (Engineers INNER JOIN
                      (Electricians INNER JOIN
                      Depletions ON Electricians.[Number] = Depletions.Electrician) ON Engineers.[Number] = Depletions.Engineer) ON 
                      Objects.[Number] = Depletions.[Object]) ON ObjectTypes.[Number] = Objects.ObjectType) ON OGZNumber.[Number] = Depletions.OGZNumber) ON 
                      OnDutys.[Number] = Depletions.OnDuty) ON Organizations.[Number] = Objects.Organization) ON 
                      ProtectionTypes.[Number] = Depletions.ProtectionType) ON Reasons2.[Number] = Depletions.FinalReason) ON 
                      Reasons.[Number] = Depletions.InitialReason) ON ShiftEngineers.[Number] = Depletions.ShiftEngineer)
            
            ORDER BY Depletions.ReportDate, Depletions.DateTimeAlarm", oleconn);

            //****************************************
            //WHERE Depletions.ReportDate >=#" + monthCalendar1.SelectionRange.Start.Month + "/" + monthCalendar1.SelectionRange.Start.Day + "/" + monthCalendar1.SelectionRange.Start.Year + @"#
            //AND Depletions.ReportDate <=#" + monthCalendar1.SelectionRange.End.Month + "/" + monthCalendar1.SelectionRange.End.Day + "/" + monthCalendar1.SelectionRange.End.Year + @"# 

            OleDbDataReader test_r1 = test1.ExecuteReader();

            toolStripProgressBar1.Value = 10;
            toolStripStatusLabel2.Text = "SQL query...";
            Refresh();

            SqlCommand imp_dep = new SqlCommand();
            imp_dep.Connection = Program.conn;

            sql_comm_str = "TRUNCATE TABLE [dbo].[uDep] ";

            imp_dep.CommandText = sql_comm_str;
            imp_dep.ExecuteNonQuery();

            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;

            string type_obj_str = "";
            string type_protect_str = "";
            string time_complete_substr = "";
            string time_complete_str = "";

            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];
                ds1[rowcount, 7] = test_r1[7];
                ds1[rowcount, 8] = test_r1[8];
                ds1[rowcount, 9] = test_r1[9];
                ds1[rowcount, 10] = test_r1[10];
                ds1[rowcount, 11] = test_r1[11];
                ds1[rowcount, 12] = test_r1[12];
                ds1[rowcount, 13] = test_r1[13];
                ds1[rowcount, 14] = test_r1[14];
                ds1[rowcount, 15] = test_r1[15];
                ds1[rowcount, 16] = test_r1[16];
                ds1[rowcount, 17] = test_r1[17];
                ds1[rowcount, 18] = test_r1[18];
                ds1[rowcount, 19] = test_r1[19];
                ds1[rowcount, 20] = test_r1[20];
                ds1[rowcount, 21] = test_r1[21];
                ds1[rowcount, 22] = test_r1[22];
                ds1[rowcount, 23] = test_r1[23];
                ds1[rowcount, 24] = test_r1[24];
                ds1[rowcount, 25] = test_r1[25];
                ds1[rowcount, 26] = test_r1[26];
                ds1[rowcount, 27] = test_r1[27];
                ds1[rowcount, 28] = test_r1[28];
                ds1[rowcount, 29] = test_r1[29];

                if (test_r1[29].ToString().Contains("(") && test_r1[29].ToString().Contains(")"))
                {
                    time_complete_substr = test_r1[29].ToString();
                    int st1 = time_complete_substr.IndexOf("(");
                    int en1 = time_complete_substr.IndexOf(")");

                    if (time_complete_substr.Contains(":"))
                    {
                        time_complete_substr = time_complete_substr.Substring(st1 + 1, en1 - (st1 + 1));
                        time_complete_str = ((DateTime)test_r1[13]).Date.ToString("dd.MM.yyyy") + " " + time_complete_substr + ":00";
                    }
                    else {
                        time_complete_str = ((DateTime)test_r1[13]).Date.ToString("dd.MM.yyyy") + " " + "00:00:00";
                    }
                }
                else
                    time_complete_str = test_r1[13].ToString();

                if (test_r1[6].ToString() == Program.list_obj_type[0]) type_obj_str = "1";
                else {
                    if (test_r1[6].ToString() == "ЛИГ") type_obj_str = "2";
                    else type_obj_str = "3";
                }
                //if (test_r1[6].ToString() == Program.list_obj_type[2]) type_obj_str = "3";

                if (test_r1[7].ToString() == Program.list_protect_type[0]) type_protect_str = "1";
                if (test_r1[7].ToString() == Program.list_protect_type[1]) type_protect_str = "2";

                sql_comm_str = "INSERT [dbo].[uDep] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[date_report] ,[date_alarm] ,[date_pass] ,[date_arrive] ,[date_complete] ,[gz] ,[gz_start] ,[comment_alarm] ,[comment_work] ,[comment_afterwork] ,[comment_general] ,[reason_start] ,[reason_end] ,[worker1] ,[worker2] ,[worker3] ,[worker4] ,[service_org_name]) VALUES ('" + test_r1[1].ToString().Replace('\'', ' ') + "', '" + test_r1[3].ToString().Replace('\'', ' ') + "', '" + test_r1[4].ToString().Replace('\'', ' ') + "','" + test_r1[5].ToString().Replace('\'', ' ') + "'," + type_obj_str + "," + type_protect_str + ",'" + ((DateTime)test_r1[8]).ToString("yyyy-dd-MM HH:mm:ss") + "','" + ((DateTime)test_r1[9]).ToString("yyyy-dd-MM HH:mm:ss") + "','" + ((DateTime)test_r1[10]).ToString("yyyy-dd-MM HH:mm:ss") + "','" + ((DateTime)test_r1[13]).ToString("yyyy-dd-MM HH:mm:ss") + "','" + (DateTime.Parse(time_complete_str)).ToString("yyyy-dd-MM HH:mm:ss") + "','" + test_r1[11].ToString().Replace('\'', ' ') + "', '" + test_r1[12].ToString().Replace('\'', ' ') + "','" + test_r1[14].ToString().Replace('\'', ' ') + "', '" + test_r1[15].ToString().Replace('\'', ' ') + "','" + test_r1[16].ToString().Replace('\'', ' ') + "', '" + test_r1[29].ToString().Replace('\'', ' ') + "','" + test_r1[18].ToString().Replace('\'', ' ') + "', '" + test_r1[20].ToString().Replace('\'', ' ') + "','" + test_r1[26].ToString().Replace('\'', ' ') + "', '" + test_r1[28].ToString().Replace('\'', ' ') + "','" + test_r1[24].ToString().Replace('\'', ' ') + "', '" + "нет (импорт)" + "','" + test_r1[22].ToString().Replace('\'', ' ') + "')";

                imp_dep.CommandText = sql_comm_str;
                imp_dep.ExecuteNonQuery();

                rowcount++;
            }

            toolStripProgressBar1.Value = 95;
            //this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);
            if (File.Exists("udep_row_count.txt"))
            {
                //
                this.uDepTableAdapter.FillBy(this.uniq1DataSet.uDep, Int32.Parse(File.ReadAllLines("udep_row_count.txt", Encoding.GetEncoding(1251))[0]));///!!!
            }
            else
                this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);

            toolStripProgressBar1.Value = 100;

        }

        private void label27_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = "";
            comboBox1.SelectedValue = comboBox2.SelectedValue = null;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {

                int ee = 0;

                string[] strerr = new string[8];
                char[] splc = new char[1];
                int ce = 0;
                splc[0] = '^';
                strerr = imp_err.Split(splc);
                //sql_comm_str = "INSERT [dbo].[uDep] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[date_report] ,[date_alarm] ,[date_pass] ,[date_arrive] ,[date_complete] ,[gz] ,[gz_start] ,[comment_alarm] ,[comment_work] ,[comment_afterwork] ,[comment_general] ,[reason_start] ,[reason_end] ,[worker1] ,[worker2] ,[worker3] ,[worker4] ,[service_org_name]) VALUES ('" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "','" + textBox6.Text + "'," + (comboBox1.SelectedIndex + 1).ToString() + "," + (comboBox2.SelectedIndex + 1).ToString() + ",'" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "','" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "','" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "','" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "','" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "','" + comboBox5.Text.ToString() + "', '" + textBox7.Text + "','" + textBox8.Text + "', '" + textBox10.Text + "','" + textBox9.Text + "', '" + textBox11.Text + "','" + comboBox3.Text.ToString() + "', '" + comboBox4.Text.ToString() + "','" + comboBox7.Text.ToString() + "', '" + comboBox8.Text.ToString() + "','" + comboBox9.Text.ToString() + "', '" + comboBox10.Text.ToString() + "','" + comboBox6.Text.ToString() + "')";

                if (textBox3.Text == strerr[1])
                {
                    strerr[1] = "";
                    ce++;
                }
                if (textBox4.Text == strerr[2])
                {
                    strerr[2] = "";
                    ce++;
                }
                if (textBox5.Text == strerr[3])
                {
                    strerr[3] = "";
                    ce++;
                }
                if (textBox6.Text == strerr[4])
                {
                    strerr[4] = "";
                    ce++;
                }
                if ((comboBox1.SelectedIndex + 1).ToString() == strerr[5])
                {
                    strerr[5] = "";
                    ce++;
                }
                if ((comboBox2.SelectedIndex + 1).ToString() == strerr[6])
                {
                    strerr[6] = "";
                    ce++;
                }



                if (textBox3.Text.Length < 2)
                {
                    MessageBox.Show("Поле %позывной% не заполнено! Не могу добавить сработку!");
                    ee++;
                }
                else
                {
                    if (textBox4.Text.Length < 2)
                    {
                        MessageBox.Show("Поле %организация% не заполнено! Не могу добавить сработку!");
                        ee++;
                    }
                    else
                    {
                        if (textBox5.Text.Length < 2)
                        {
                            MessageBox.Show("Поле %объект% не заполнено! Не могу добавить сработку!");
                            ee++;
                        }
                        else
                        {
                            if (textBox6.Text.Length < 2)
                            {
                                MessageBox.Show("Поле %адрес% не заполнено! Не могу добавить сработку!");
                                ee++;
                            }
                            else
                            {
                                if (comboBox1.SelectedIndex < 0)
                                {
                                    MessageBox.Show("Поле %тип объекта% не заполнено! Не могу добавить сработку!");
                                    ee++;
                                }
                                else
                                {
                                    if (comboBox2.SelectedIndex < 0)
                                    {
                                        MessageBox.Show("Поле %тип охраны% не заполнено! Не могу добавить сработку!");
                                        ee++;
                                    }
                                    else
                                    {
                                        if (comboBox3.SelectedIndex < 0)
                                        {
                                            MessageBox.Show("Поле %начальная причина% не заполнено! Не могу добавить сработку!");
                                            ee++;
                                        }
                                        else
                                        {
                                            if (comboBox4.SelectedIndex < 0)
                                            {
                                                MessageBox.Show("Поле %конечная причина% не заполнено! Не могу добавить сработку!");
                                                ee++;
                                            }
                                            else
                                            {
                                                if (comboBox7.SelectedIndex < 0)
                                                {
                                                    MessageBox.Show("Поле %дежурный ПУ% не заполнено! Не могу добавить сработку!");
                                                    ee++;
                                                }
                                                else
                                                {
                                                    if (comboBox8.SelectedIndex < 0)
                                                    {
                                                        MessageBox.Show("Поле %дежурный офицер% не заполнено! Не могу добавить сработку!");
                                                        ee++;
                                                    }
                                                    else
                                                    {
                                                        if (dateTimePicker2.Value.AddHours(dateTimePicker3.Value.Hour).AddMinutes(dateTimePicker3.Value.Minute).AddSeconds(dateTimePicker3.Value.Second) >= dateTimePicker5.Value.AddHours(dateTimePicker4.Value.Hour).AddMinutes(dateTimePicker4.Value.Minute).AddSeconds(dateTimePicker4.Value.Second) && comboBox3.Text != "Заявка")
                                                        {
                                                            MessageBox.Show("Хронология событий нарушена! Событие передача раньше сработки! Не могу добавить сработку!");
                                                            ee++;
                                                        }
                                                        else
                                                        {
                                                            if (dateTimePicker5.Value.AddHours(dateTimePicker4.Value.Hour).AddMinutes(dateTimePicker4.Value.Minute).AddSeconds(dateTimePicker4.Value.Second) >= dateTimePicker7.Value.AddHours(dateTimePicker6.Value.Hour).AddMinutes(dateTimePicker6.Value.Minute).AddSeconds(dateTimePicker6.Value.Second) && comboBox3.Text != "Заявка")
                                                            {
                                                                MessageBox.Show("Хронология событий нарушена! Время прибытия раньше передачи! Не могу добавить сработку!");
                                                                ee++;
                                                            }
                                                            else
                                                            {
                                                                if (dateTimePicker7.Value.AddHours(dateTimePicker6.Value.Hour).AddMinutes(dateTimePicker6.Value.Minute).AddSeconds(dateTimePicker6.Value.Second) >= dateTimePicker9.Value.AddHours(dateTimePicker8.Value.Hour).AddMinutes(dateTimePicker8.Value.Minute).AddSeconds(dateTimePicker8.Value.Second) && comboBox3.Text != "Заявка")
                                                                {
                                                                    MessageBox.Show("Хронология событий нарушена! Событие %осмотр завершен% раньше времени прибытия! Не могу добавить сработку!");
                                                                    ee++;
                                                                }
                                                                else
                                                                {
                                                                    if (textBox13.Text.Length < 2)
                                                                    {
                                                                        MessageBox.Show("Поле %Сработавшее оборудование-> Название датчика% не заполнено! Не могу добавить сработку!");
                                                                        ee++;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (textBox14.Text.Length < 2)
                                                                        {
                                                                            MessageBox.Show("Поле %Погодные условия-> Температура % не заполнено! Не могу добавить сработку!");
                                                                            ee++;
                                                                        }
                                                                        else
                                                                        {
                                                                            if (comboBox2.SelectedIndex < 0)
                                                                            {
                                                                                MessageBox.Show("Поле %Погодные условия->Доп. сведения% не заполнено! Не могу добавить сработку!");
                                                                                ee++;
                                                                            }
                                                                            else
                                                                            {
                                                                                if (textBox18.Text.Length < 2 && ce < 6)
                                                                                {
                                                                                    MessageBox.Show("Поле %Коментарий к изменениям% Обязательно для заполнения! Не могу добавить сработку!");
                                                                                    ee++;
                                                                                }

                                                                            }
                                                                        }

                                                                    }
                                                                }


                                                            }


                                                        }


                                                    }


                                                }


                                            }


                                        }

                                    }


                                }
                            }
                        }
                    }
                }
                ///****************************************************************//////////////////////////////////////////////////////////////////////////////////////////

                if (imp_err == null)
                {
                    MessageBox.Show("Блок данных не выбран из дислокации!!!", "pcn_box", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    imp_err = "^^^^^^";
                }

                if (ee == 0)
                {
                    if (ce < 6)
                        File.AppendAllText("transact_" + DateTime.Now.ToString("yyyy-dd-MM") + ".txt", "INS&" + textBox3.Text + "&" + textBox4.Text + "&" + textBox5.Text + "&" + textBox6.Text + "&" + (comboBox1.SelectedIndex + 1).ToString() + "&" + (comboBox2.SelectedIndex + 1).ToString() + "&" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "&" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "&" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "&" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "&" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "&" + comboBox5.Text.ToString() + "&" + textBox7.Text + "&" + textBox8.Text + "&" + textBox10.Text + "&" + textBox9.Text + "&" + textBox11.Text + "&" + comboBox3.Text.ToString() + "&" + comboBox4.Text.ToString() + "&" + comboBox7.Text.ToString() + "&" + comboBox8.Text.ToString() + "&" + comboBox9.Text.ToString() + "&" + comboBox10.Text.ToString() + "&" + comboBox6.Text.ToString() + "&" + strerr[0] + "^" + strerr[1] + "^" + strerr[2] + "^" + strerr[3] + "^" + strerr[4] + "^" + strerr[5] + "^" + strerr[6] + "^" + "&" + textBox13.Text + "&" + textBox14.Text + "&" + comboBox11.Text.ToString() + "&" + textBox16.Text + "&" + textBox17.Text + "&" + textBox18.Text + "&" + textBox20.Text + "&" + textBox19.Text + "\n", Encoding.GetEncoding(1251));
                    //sql_comm_str = "INSERT [dbo].[uDep] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[date_report] ,[date_alarm] ,[date_pass] ,[date_arrive] ,[date_complete] ,[gz] ,[gz_start] ,[comment_alarm] ,[comment_work] ,[comment_afterwork] ,[comment_general] ,[reason_start] ,[reason_end] ,[worker1] ,[worker2] ,[worker3] ,[worker4] ,[service_org_name], [import_errors], [e_ment], [temper], [weather], [dogovor1], [debitorska1], [comment_edit], [average_month], [last_month]) VALUES ('" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "','" + textBox6.Text + "'," + (comboBox1.SelectedIndex + 1).ToString() + "," + (comboBox2.SelectedIndex + 1).ToString() + ",'" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "','" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "','" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "','" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "','" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "','" + comboBox5.Text.ToString() + "', '" + textBox7.Text + "','" + textBox8.Text + "', '" + textBox10.Text + "','" + textBox9.Text + "', '" + textBox11.Text + "','" + comboBox3.Text.ToString() + "', '" + comboBox4.Text.ToString() + "','" + comboBox7.Text.ToString() + "', '" + comboBox8.Text.ToString() + "','" + comboBox9.Text.ToString() + "', '" + comboBox10.Text.ToString() + "','" + comboBox6.Text.ToString() + "','" + strerr[0] + "^" + strerr[1] + "^" + strerr[2] + "^" + strerr[3] + "^" + strerr[4] + "^" + strerr[5] + "^" + strerr[6] + "^" + "','" + textBox13.Text + "','" + textBox14.Text + "','" + comboBox11.Text.ToString() + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox18.Text + "','" + textBox20.Text + "','" + textBox19.Text + "')";
                    else
                        File.AppendAllText("transact_" + DateTime.Now.ToString("yyyy-dd-MM") + ".txt", "INS&" + textBox3.Text + "&" + textBox4.Text + "&" + textBox5.Text + "&" + textBox6.Text + "&" + (comboBox1.SelectedIndex + 1).ToString() + "&" + (comboBox2.SelectedIndex + 1).ToString() + "&" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "&" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "&" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "&" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "&" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "&" + comboBox5.Text.ToString() + "&" + textBox7.Text + "&" + textBox8.Text + "&" + textBox10.Text + "&" + textBox9.Text + "&" + textBox11.Text + "&" + comboBox3.Text.ToString() + "&" + comboBox4.Text.ToString() + "&" + comboBox7.Text.ToString() + "&" + comboBox8.Text.ToString() + "&" + comboBox9.Text.ToString() + "&" + comboBox10.Text.ToString() + "&" + comboBox6.Text.ToString() + "&" + textBox13.Text + "&" + textBox14.Text + "&" + comboBox11.Text.ToString() + "&" + textBox16.Text + "&" + textBox17.Text + "&" + textBox20.Text + "&" + textBox19.Text + "\n", Encoding.GetEncoding(1251));
                    //sql_comm_str = "INSERT [dbo].[uDep] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[date_report] ,[date_alarm] ,[date_pass] ,[date_arrive] ,[date_complete] ,[gz] ,[gz_start] ,[comment_alarm] ,[comment_work] ,[comment_afterwork] ,[comment_general] ,[reason_start] ,[reason_end] ,[worker1] ,[worker2] ,[worker3] ,[worker4] ,[service_org_name], [e_ment], [temper], [weather], [dogovor1], [debitorska1], [average_month], [last_month]) VALUES ('" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "','" + textBox6.Text + "'," + (comboBox1.SelectedIndex + 1).ToString() + "," + (comboBox2.SelectedIndex + 1).ToString() + ",'" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "','" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "','" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "','" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "','" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "','" + comboBox5.Text.ToString() + "', '" + textBox7.Text + "','" + textBox8.Text + "', '" + textBox10.Text + "','" + textBox9.Text + "', '" + textBox11.Text + "','" + comboBox3.Text.ToString() + "', '" + comboBox4.Text.ToString() + "','" + comboBox7.Text.ToString() + "', '" + comboBox8.Text.ToString() + "','" + comboBox9.Text.ToString() + "', '" + comboBox10.Text.ToString() + "','" + comboBox6.Text.ToString() + "','" + textBox13.Text + "','" + textBox14.Text + "','" + comboBox11.Text.ToString() + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox20.Text + "','" + textBox19.Text + "')";

                    //File.AppendAllText("transact_"+DateTime.Now.ToString("yyyy-dd-MM")+".txt" ,"INS&"+  , Encoding.GetEncoding(1251));

                    string sql_comm_str = "";
                    SqlCommand imp_dep = new SqlCommand();
                    imp_dep.Connection = Program.conn;


                    if (ce < 6)
                        sql_comm_str = "INSERT [dbo].[uDep] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[date_report] ,[date_alarm] ,[date_pass] ,[date_arrive] ,[date_complete] ,[gz] ,[gz_start] ,[comment_alarm] ,[comment_work] ,[comment_afterwork] ,[comment_general] ,[reason_start] ,[reason_end] ,[worker1] ,[worker2] ,[worker3] ,[worker4] ,[service_org_name], [import_errors], [e_ment], [temper], [weather], [dogovor1], [debitorska1], [comment_edit], [average_month], [last_month]) VALUES ('" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "','" + textBox6.Text + "'," + (comboBox1.SelectedIndex + 1).ToString() + "," + (comboBox2.SelectedIndex + 1).ToString() + ",'" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "','" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "','" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "','" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "','" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "','" + comboBox5.Text.ToString() + "', '" + textBox7.Text + "','" + textBox8.Text + "', '" + textBox10.Text + "','" + textBox9.Text + "', '" + textBox11.Text + "','" + comboBox3.Text.ToString() + "', '" + comboBox4.Text.ToString() + "','" + comboBox7.Text.ToString() + "', '" + comboBox8.Text.ToString() + "','" + comboBox9.Text.ToString() + "', '" + comboBox10.Text.ToString() + "','" + comboBox6.Text.ToString() + "','" + strerr[0] + "^" + strerr[1] + "^" + strerr[2] + "^" + strerr[3] + "^" + strerr[4] + "^" + strerr[5] + "^" + strerr[6] + "^" + "','" + textBox13.Text + "','" + textBox14.Text + "','" + comboBox11.Text.ToString() + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox18.Text + "','" + textBox20.Text + "','" + textBox19.Text + "')";
                    else
                        sql_comm_str = "INSERT [dbo].[uDep] ([poz] ,[obj_org] ,[obj_name] ,[obj_adr] ,[obj_type] ,[type_protect] ,[date_report] ,[date_alarm] ,[date_pass] ,[date_arrive] ,[date_complete] ,[gz] ,[gz_start] ,[comment_alarm] ,[comment_work] ,[comment_afterwork] ,[comment_general] ,[reason_start] ,[reason_end] ,[worker1] ,[worker2] ,[worker3] ,[worker4] ,[service_org_name], [e_ment], [temper], [weather], [dogovor1], [debitorska1], [average_month], [last_month]) VALUES ('" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "','" + textBox6.Text + "'," + (comboBox1.SelectedIndex + 1).ToString() + "," + (comboBox2.SelectedIndex + 1).ToString() + ",'" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "','" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "','" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "','" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "','" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "','" + comboBox5.Text.ToString() + "', '" + textBox7.Text + "','" + textBox8.Text + "', '" + textBox10.Text + "','" + textBox9.Text + "', '" + textBox11.Text + "','" + comboBox3.Text.ToString() + "', '" + comboBox4.Text.ToString() + "','" + comboBox7.Text.ToString() + "', '" + comboBox8.Text.ToString() + "','" + comboBox9.Text.ToString() + "', '" + comboBox10.Text.ToString() + "','" + comboBox6.Text.ToString() + "','" + textBox13.Text + "','" + textBox14.Text + "','" + comboBox11.Text.ToString() + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox20.Text + "','" + textBox19.Text + "')";
                    imp_dep.CommandText = sql_comm_str;
                    imp_dep.ExecuteNonQuery();
                    Clear_Fields();

                }
            }
            else MessageBox.Show("Выбрано другое дествие с записью данных!");

            //this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);
            if (File.Exists("udep_row_count.txt"))
            {
                //
                this.uDepTableAdapter.FillBy(this.uniq1DataSet.uDep, Int32.Parse(File.ReadAllLines("udep_row_count.txt", Encoding.GetEncoding(1251))[0]));///!!!
            }
            else
                this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);


        }

        private void стандартныйОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            Process.Start("st_report.exe");
        }

        private void вЫХОДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && radioButton2.Checked)
            {
                label27.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                textBox3.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox4.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                textBox5.Text = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                textBox6.Text = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "1")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[0]);///*****************
                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "2")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[1]);///*****************
                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "3")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[2]);///*****************
                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "4")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[3]);///*****************
                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "5")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[4]);///*****************
                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "6")
                    comboBox1.SelectedIndex = comboBox1.FindString(Program.list_obj_type[5]);///*****************


                if (dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString() == "1")
                    comboBox2.SelectedIndex = comboBox2.FindString(Program.list_protect_type[0]);///*****************
                if (dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString() == "2")
                    comboBox2.SelectedIndex = comboBox2.FindString(Program.list_protect_type[1]);///*****************

                /////////////////////////////////////////////////////////////////////////////////////////////////////

                dateTimePicker1.Value = DateTime.Parse(dataGridView2.Rows[e.RowIndex].Cells[7].Value.ToString());
                dateTimePicker2.Value = dateTimePicker3.Value = DateTime.Parse(dataGridView2.Rows[e.RowIndex].Cells[8].Value.ToString());
                dateTimePicker5.Value = dateTimePicker4.Value = DateTime.Parse(dataGridView2.Rows[e.RowIndex].Cells[9].Value.ToString());
                dateTimePicker7.Value = dateTimePicker6.Value = DateTime.Parse(dataGridView2.Rows[e.RowIndex].Cells[10].Value.ToString());
                dateTimePicker9.Value = dateTimePicker8.Value = DateTime.Parse(dataGridView2.Rows[e.RowIndex].Cells[11].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[29].Value.ToString() != "")
                    dateTimePicker10.Value = DateTime.Parse(dataGridView2.Rows[e.RowIndex].Cells[29].Value.ToString());
                else
                    dateTimePicker10.Value = DateTime.Now;

                //comboBox5.Text.ToString() + "', '" + textBox7.Text + "','" + textBox8.Text + "', '" + textBox10.Text + "','" + textBox9.Text + "', '" + textBox11.Text + "','" + comboBox3.Text.ToString() + "', '" + comboBox4.Text.ToString() + "','" + comboBox7.Text.ToString() + "', '" + comboBox8.Text.ToString() + "','" + comboBox9.Text.ToString() + "', '" + comboBox10.Text.ToString() + "','" + comboBox6.Text.ToString()
                if (dataGridView2.Rows[e.RowIndex].Cells[12].Value.ToString() != "")
                    comboBox5.SelectedIndex = comboBox5.FindString(dataGridView2.Rows[e.RowIndex].Cells[12].Value.ToString());
                textBox7.Text = dataGridView2.Rows[e.RowIndex].Cells[13].Value.ToString();
                textBox8.Text = dataGridView2.Rows[e.RowIndex].Cells[14].Value.ToString();
                textBox10.Text = dataGridView2.Rows[e.RowIndex].Cells[15].Value.ToString();
                textBox9.Text = dataGridView2.Rows[e.RowIndex].Cells[16].Value.ToString();
                textBox11.Text = dataGridView2.Rows[e.RowIndex].Cells[17].Value.ToString();
                if (dataGridView2.Rows[e.RowIndex].Cells[18].Value.ToString() != "")
                    comboBox3.SelectedIndex = comboBox3.FindString(dataGridView2.Rows[e.RowIndex].Cells[18].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[19].Value.ToString() != "")
                    comboBox4.SelectedIndex = comboBox4.FindString(dataGridView2.Rows[e.RowIndex].Cells[19].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[20].Value.ToString() != "")
                    comboBox7.SelectedIndex = comboBox7.FindString(dataGridView2.Rows[e.RowIndex].Cells[20].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[21].Value.ToString() != "")
                    comboBox8.SelectedIndex = comboBox8.FindString(dataGridView2.Rows[e.RowIndex].Cells[21].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[22].Value.ToString() != "")
                    comboBox9.SelectedIndex = comboBox9.FindString(dataGridView2.Rows[e.RowIndex].Cells[22].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[23].Value.ToString() != "")
                    comboBox10.SelectedIndex = comboBox10.FindString(dataGridView2.Rows[e.RowIndex].Cells[23].Value.ToString());
                if (dataGridView2.Rows[e.RowIndex].Cells[24].Value.ToString() != "")
                    comboBox6.SelectedIndex = comboBox6.FindString(dataGridView2.Rows[e.RowIndex].Cells[24].Value.ToString());

                textBox13.Text = dataGridView2.Rows[e.RowIndex].Cells[25].Value.ToString();
                textBox14.Text = dataGridView2.Rows[e.RowIndex].Cells[26].Value.ToString();
                textBox16.Text = dataGridView2.Rows[e.RowIndex].Cells[30].Value.ToString();
                //textBox17.Text = dataGridView2.Rows[e.RowIndex].Cells[30].Value.ToString();

                textBox18.Text = dataGridView2.Rows[e.RowIndex].Cells[31].Value.ToString();
                textBox20.Text = dataGridView2.Rows[e.RowIndex].Cells[32].Value.ToString();
                textBox19.Text = dataGridView2.Rows[e.RowIndex].Cells[33].Value.ToString();

                if (dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString() == "1")
                {
                    groupBox11.Enabled = true;
                }
                else
                    groupBox11.Enabled = false;

                if (dataGridView2.Rows[e.RowIndex].Cells[27].Value.ToString() != "")
                    comboBox11.SelectedIndex = comboBox11.FindString(dataGridView2.Rows[e.RowIndex].Cells[27].Value.ToString());

            }

        }

        public void Clear_Fields()
        {
            button1.Enabled = true;
            textBox13.Text = textBox14.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = textBox7.Text = textBox8.Text = textBox10.Text = textBox9.Text = textBox11.Text = textBox16.Text = textBox17.Text = textBox18.Text = textBox19.Text = textBox20.Text = "";
            comboBox1.SelectedIndex = comboBox2.SelectedIndex = comboBox3.SelectedIndex = comboBox4.SelectedIndex = comboBox5.SelectedIndex = comboBox6.SelectedIndex = comboBox7.SelectedIndex = comboBox8.SelectedIndex = comboBox9.SelectedIndex = comboBox10.SelectedIndex = comboBox11.SelectedIndex = -1;



            if (DateTime.Now.Hour < rep_time)
            {
                DateTime r = new DateTime();
                r = DateTime.Now.AddDays(-1);

                dateTimePicker1.Value = r;
            }
            else
                dateTimePicker1.Value = DateTime.Now;


            dateTimePicker2.Value = dateTimePicker3.Value = DateTime.Now;
            dateTimePicker5.Value = dateTimePicker4.Value = DateTime.Now;
            dateTimePicker7.Value = dateTimePicker6.Value = DateTime.Now;
            dateTimePicker9.Value = dateTimePicker8.Value = DateTime.Now;
            dateTimePicker10.Value = DateTime.Now;

            label27.Text = "0";
            label32.Text = "0";

            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            checkBox9.Checked = false;

            groupBox11.Enabled = false;

        }
        private void button2_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                dataGridView2.DefaultCellStyle.BackColor = Color.LightGray;
                dataGridView1.DefaultCellStyle.BackColor = Color.White;
                label38.Visible = false;
                dateTimePicker10.Visible = false;
            }
            else
            {
                dataGridView1.DefaultCellStyle.BackColor = Color.LightGray;
                dataGridView2.DefaultCellStyle.BackColor = Color.White;
                label38.Visible = true;
                dateTimePicker10.Visible = true;

            }
            Clear_Fields();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (radioButton2.Checked)
            {
                if (MessageBox.Show("Удалить сработку #" + label27.Text + " ?", "pcn_box", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {

                    File.AppendAllText("transact_" + DateTime.Now.ToString("yyyy-dd-MM") + ".txt", "DEL&" + label27.Text + "\n", Encoding.GetEncoding(1251));

                    string sql_comm_str = "";
                    SqlCommand imp_dep = new SqlCommand();
                    imp_dep.Connection = Program.conn;

                    sql_comm_str = "DELETE FROM [dbo].[uDep] WHERE idd=" + label27.Text;
                    imp_dep.CommandText = sql_comm_str;
                    imp_dep.ExecuteNonQuery();
                }
                else
                    MessageBox.Show("Сработка не удалена!", "pcn_box", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            else MessageBox.Show("Выбрано другое дествие с записью данных!");

            //this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);
            if (File.Exists("udep_row_count.txt"))
            {
                //
                this.uDepTableAdapter.FillBy(this.uniq1DataSet.uDep, Int32.Parse(File.ReadAllLines("udep_row_count.txt", Encoding.GetEncoding(1251))[0]));///!!!
            }
            else
                this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);

            Clear_Fields();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //
            if (radioButton2.Checked)
            {
                File.AppendAllText("transact_" + DateTime.Now.ToString("yyyy-dd-MM") + ".txt", "UPD&" + textBox3.Text + "&" + textBox4.Text + "&" + textBox5.Text + "&" + textBox6.Text + "&" + (comboBox1.SelectedIndex + 1).ToString() + "&" + (comboBox2.SelectedIndex + 1).ToString() + "&" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "&" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "&" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "&" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "&" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "&" + comboBox5.Text.ToString() + "&" + textBox7.Text + "&" + textBox8.Text + "&" + textBox10.Text + "&" + textBox9.Text + "&" + textBox11.Text + "&" + comboBox3.Text.ToString() + "&" + comboBox4.Text.ToString() + "&" + comboBox7.Text.ToString() + "&" + comboBox8.Text.ToString() + "&" + comboBox9.Text.ToString() + "&" + comboBox10.Text.ToString() + "&" + comboBox6.Text.ToString() + "&" + textBox13.Text + "&" + textBox14.Text + "&" + comboBox11.Text.ToString() + "&" + dateTimePicker10.Value.ToString("yyyy-dd-MM") + "&" + textBox18.Text + "&" + textBox19.Text + "&" + label27.Text + "\n", Encoding.GetEncoding(1251));

                string sql_comm_str = "";
                SqlCommand imp_dep = new SqlCommand();
                imp_dep.Connection = Program.conn;
                sql_comm_str = "UPDATE [dbo].[uDep] SET [poz]='" + textBox3.Text + "', [obj_org]='" + textBox4.Text + "', [obj_name]='" + textBox5.Text + "', [obj_adr]='" + textBox6.Text + "', [obj_type]=" + (comboBox1.SelectedIndex + 1).ToString() + ", [type_protect]=" + (comboBox2.SelectedIndex + 1).ToString() + ", [date_report]='" + dateTimePicker1.Value.ToString("yyyy-dd-MM") + "', [date_alarm]='" + dateTimePicker2.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker3.Value.ToString("HH:mm:ss") + "', [date_pass]='" + dateTimePicker5.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker4.Value.ToString("HH:mm:ss") + "', [date_arrive]='" + dateTimePicker7.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker6.Value.ToString("HH:mm:ss") + "', [date_complete]='" + dateTimePicker9.Value.ToString("yyyy-dd-MM") + " " + dateTimePicker8.Value.ToString("HH:mm:ss") + "', [gz]='" + comboBox5.Text.ToString() + "', [gz_start]='" + textBox7.Text + "', [comment_alarm]='" + textBox8.Text + "', [comment_work]='" + textBox10.Text + "', [comment_afterwork]='" + textBox9.Text + "', [comment_general]='" + textBox11.Text + "', [reason_start]='" + comboBox3.Text.ToString() + "', [reason_end]='" + comboBox4.Text.ToString() + "', [worker1]='" + comboBox7.Text.ToString() + "', [worker2]='" + comboBox8.Text.ToString() + "', [worker3]='" + comboBox9.Text.ToString() + "', [worker4]='" + comboBox10.Text.ToString() + "', [service_org_name]='" + comboBox6.Text.ToString() + "', [e_ment]='" + textBox13.Text + "', [temper]='" + textBox14.Text + "', [weather]='" + comboBox11.Text.ToString() + "', [date_afterwork]='" + dateTimePicker10.Value.ToString("yyyy-dd-MM") + "', [comment_edit]='" + textBox18.Text + "', [last_month]='" + textBox19.Text + "' WHERE idd=" + label27.Text; // + "," + 
                imp_dep.CommandText = sql_comm_str;
                imp_dep.ExecuteNonQuery();
            }
            else MessageBox.Show("Выбрано другое дествие с записью данных!");


            //this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);
            if (File.Exists("udep_row_count.txt"))
            {
                //
                this.uDepTableAdapter.FillBy(this.uniq1DataSet.uDep, Int32.Parse(File.ReadAllLines("udep_row_count.txt", Encoding.GetEncoding(1251))[0]));///!!!
            }
            else
                this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);

            Clear_Fields();
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked)
            {
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                textBox5.Enabled = true;
                textBox6.Enabled = true;
                textBox16.Enabled = true;
                textBox17.Enabled = true;
                textBox18.Enabled = true;
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;

            }
            else
            {
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                textBox16.Enabled = false;
                textBox17.Enabled = false;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;

            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (textBox12.Text == "uniqsight")
            {
                panel1.Visible = false;
                textBox12.Text = "";
                импортИзБДDepletionToolStripMenuItem.Enabled = true;
                пользователиToolStripMenuItem.Enabled = true;
                label33.Visible = true;
                button1.Visible = true;
                button3.Visible = true;
                button4.Visible = true;

            }
            if (textBox12.Text == "1912")
            {
                panel1.Visible = false;
                textBox12.Text = "";
                label33.Visible = true;
                button1.Visible = true;
                button3.Visible = true;
                button4.Visible = true;

            }

            if (textBox12.Text == "9998")
            {
                panel1.Visible = false;
                textBox12.Text = "";
                label33.Visible = true;
                button1.Visible = true;
                button3.Visible = true;
                button4.Visible = true;

                Program.sql_insert_deblo = true;


            }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void регистрацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            label33.Visible = false;
            button1.Visible = false;
            button3.Visible = false;
            button4.Visible = false;


        }

        private void отчетЗаПериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            Process.Start("vd_report.exe");
        }

        private void textBox10_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Program.comments = 1;
            Form2 frm2 = new Form2();
            frm2.ShowDialog();
            textBox10.Text = Program.fast_str;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_DoubleClick(object sender, EventArgs e)
        {

        }

        private void textBox8_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Program.comments = 2;
            Form2 frm2 = new Form2();
            frm2.ShowDialog();
            textBox8.Text = Program.fast_str;
        }

        private void обновитьТаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);
            if (File.Exists("udep_row_count.txt"))
            {
                //
                this.uDepTableAdapter.FillBy(this.uniq1DataSet.uDep, Int32.Parse(File.ReadAllLines("udep_row_count.txt", Encoding.GetEncoding(1251))[0]));///!!!
            }
            else
                this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);

        }

        private void отчетПоОшибкамДислокацийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            Process.Start("ed_report.exe");

        }

        private void помощьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            MessageBox.Show("Я написал программу! (с) uniq\ne-mail: uniq4ever@mail.ru\ntel: +7978-812-66-69 ", "pcn_b0x 2o18 version: " + File.ReadAllLines("ver.txt", Encoding.GetEncoding(1251))[0], MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            /**/
            if (textBox14.Text != "" && textBox14.Text != "-")
            {
                try
                {
                    Int32.Parse(textBox14.Text);
                }
                catch
                {
                    textBox14.Text = "";
                    MessageBox.Show("Температура указана не верно...");

                }
            }

        }


        public void Load_fucking_deblo()
        {

            object[,] arr14 = new object[10001, 1];
            object[,] arr15 = new object[10001, 1];
            object[,] arr16 = new object[10001, 1];
            object[,] arr17 = new object[10001, 1];
            object[,] arr18 = new object[10001, 1];
            object[,] arr19 = new object[10001, 1];

            object[,] arr1 = new object[10001, 1];
            object[,] arr2 = new object[10001, 1];
            object[,] arr3 = new object[10001, 1];
            object[,] arr4 = new object[10001, 1];
            object[,] arr5 = new object[10001, 1];
            object[,] arr6 = new object[10001, 1];

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.ShowDialog(this);
            //MessageBox.Show("Путь к файлу: " + fdlg.FileName);

            toolStripProgressBar1.Value = 5;

            if (File.Exists(fdlg.FileName) && fdlg.FileName.Contains("\\\\") == false)
            {
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut2 = ObjExcel.Workbooks.Open(fdlg.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOut2;
                ObjWorkSheetOut2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut2.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range3 = ObjWorkSheetOut2.get_Range("A" + "1", "A" + "10000");
                arr14 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("S" + "1", "S" + "10000");
                arr15 = (object[,])range3.Value2;
                //AI AJ AK AL
                range3 = ObjWorkSheetOut2.get_Range("AI" + "1", "AI" + "10000");
                arr16 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("AJ" + "1", "AJ" + "10000");
                arr17 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("AK" + "1", "AK" + "10000");
                arr18 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("AL" + "1", "AL" + "10000");
                arr19 = (object[,])range3.Value2;


                toolStripProgressBar1.Value = 10;
                /*
                if (File.Exists("E:\\evan_handler.txt"))
                {
                    File.Delete("E:\\evan_handler.txt");
                }
                */
                groupBox10.Visible = true;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = dataGridView1.RowCount;

                for (int r = 1; r < dataGridView1.RowCount; r++)
                {
                    progressBar1.Value = r;
                    if (dataGridView1.Rows[r].Cells[2].Value.ToString() != null && (dataGridView1.Rows[r].Cells[5].Value.ToString() == "2" || dataGridView1.Rows[r].Cells[5].Value.ToString() == "3"))
                    {
                        string temp_str2 = dataGridView1.Rows[r].Cells[2].Value.ToString().Trim();
                        if (temp_str2.Contains(' '))
                        {
                            //Console.Write("ws_in= " + temp_str2 + " r=" + r.ToString() + "; ");
                            string str_wo_ws = "";
                            for (int s = 0; s < temp_str2.Length; s++)
                            {
                                if (temp_str2[s] == '(')
                                    break;
                                if (temp_str2[s] != ' ')
                                    str_wo_ws += temp_str2[s];


                            }
                            arr1[r, 0] = str_wo_ws;

                            arr2[r, 0] = dataGridView1.Rows[r].Cells[0].Value.ToString().Trim();
                            arr3[r, 0] = dataGridView1.Rows[r].Cells[1].Value.ToString().Trim();
                            arr4[r, 0] = dataGridView1.Rows[r].Cells[9].Value.ToString().Trim();
                            arr5[r, 0] = dataGridView1.Rows[r].Cells[8].Value.ToString().Trim();

                            arr6[r, 0] = dataGridView1.Rows[r].Cells[11].Value.ToString().Trim();

                            //File.AppendAllText("E:\\evan_handler.txt", DateTime.Now + " \t" + arr1[r,0] + " \t" + arr2[r, 0] + " " + arr3[r, 0] + "\n", Encoding.GetEncoding(1251));
                        }

                    }
                }

                toolStripProgressBar1.Value = 15;

                progressBar1.Minimum = 0;
                progressBar1.Maximum = 10000;


                for (int r = 12; r < 10001; r++)
                {
                    progressBar1.Value = r;
                    if (arr14[r, 1] != null)
                    {
                        string temp_str2 = arr14[r, 1].ToString().Trim();

                        /*/**/

                        string[] str2_arr = new string[20];
                        char[] splc = new char[2];
                        //int ce = 0;
                        splc[0] = ' ';
                        splc[1] = '№';
                        str2_arr = temp_str2.Split(splc);

                        for (int s = 0; s < str2_arr.Length; s++)
                        {
                            if (str2_arr[s].Contains('-') && str2_arr[s].Contains('/'))
                            {
                                temp_str2 = str2_arr[s];
                                arr14[r, 1] = temp_str2;
                                break;
                            }

                        }

                        /**/



                        if (temp_str2.Contains(' '))
                        {
                            //Console.Write("ws_in= " + temp_str2 + " r=" + r.ToString() + "; ");
                            string str_wo_ws = "";
                            for (int s = 0; s < temp_str2.Length; s++)
                            {
                                if (temp_str2[s] != ' ')
                                    str_wo_ws += temp_str2[s];

                            }
                            arr14[r, 1] = str_wo_ws;

                            //Console.WriteLine("wo_ws= " + str_wo_ws + "; ");

                            if (arr15[r, 1] == null)
                                arr15[r, 1] = (string)"";

                            //File.AppendAllText("E:\\evan_handler.txt", DateTime.Now + " \t" + arr14[r, 1].ToString() + " \t" + arr15[r, 1].ToString() + "\n", Encoding.GetEncoding(1251));
                        }
                    }
                }

                toolStripProgressBar1.Value = 40;

                progressBar1.Minimum = 0;
                progressBar1.Maximum = 10000;

                progressBar2.Minimum = 0;
                progressBar2.Maximum = dataGridView1.RowCount;

                for (int r = 12; r < 10001; r++)
                {
                    progressBar1.Value = r;

                    if (arr14[r, 1] != null)
                    {
                        if (arr14[r, 1].ToString().ToLower().Contains("-") && arr14[r, 1].ToString().ToLower().Contains('/') && arr16[r - 1, 1] != null)
                        {
                            arr16[r, 1] = "_!";
                        }

                        for (int j = 1; j < dataGridView1.RowCount; j++)
                        {
                            //progressBar2.Value = j;
                            if (arr1[j, 0] != null)
                            {
                                if (arr14[r, 1].ToString().Trim().ToLower() == arr1[j, 0].ToString().Trim().ToLower())
                                {
                                    arr16[r, 1] = arr3[j, 0].ToString();
                                    arr17[r, 1] = arr4[j, 0].ToString();
                                    arr19[r, 1] = arr6[j, 0].ToString();

                                    //File.AppendAllText("E:\\evan_handler.txt", DateTime.Now + " \t" + r.ToString() + " \t" + arr14[r, 1].ToString() + " \t" + arr16[r, 1].ToString() + "\n", Encoding.GetEncoding(1251));
                                    break;
                                }

                                if (arr14[r, 1].ToString().Trim().ToLower() == arr5[j, 0].ToString().Trim().ToLower())
                                {
                                    arr18[r, 1] = arr3[j, 0].ToString();
                                    arr17[r, 1] = arr4[j, 0].ToString();
                                    arr19[r, 1] = arr6[j, 0].ToString();

                                    if (Program.sql_insert_deblo)
                                    {
                                        string sql_comm_str = "";
                                        SqlCommand debiki = new SqlCommand();
                                        debiki.Connection = Program.conn;
                                        sql_comm_str = "UPDATE [dbo].[uDisl] SET [debitorska]='" + arr15[r, 1] + "' WHERE id_imp = " + arr2[j, 0].ToString();
                                        debiki.CommandText = sql_comm_str;
                                        debiki.ExecuteNonQuery();
                                        debiki.Dispose();
                                    }



                                    //File.AppendAllText("E:\\evan_handler.txt", DateTime.Now + " \t" + r.ToString() + " \t" + arr14[r, 1].ToString() + " \t" + arr16[r, 1].ToString() + "\n", Encoding.GetEncoding(1251));
                                    break;
                                }

                            }
                        }

                    }
                }

                toolStripProgressBar1.Value = 80;

                arr16[196, 1] = "По ФИО";
                arr18[196, 1] = "По № Договора";
                arr17[196, 1] = "Контактные телефоны";
                arr19[196, 1] = "Ответственный";
                //AI AJ AK AL
                ObjWorkSheetOut2.get_Range("AI" + "1", "AI" + "10000").Value2 = arr16;
                ObjWorkSheetOut2.get_Range("AJ" + "1", "AJ" + "10000").Value2 = arr18;
                ObjWorkSheetOut2.get_Range("AK" + "1", "AK" + "10000").Value2 = arr17;
                ObjWorkSheetOut2.get_Range("AL" + "1", "AL" + "10000").Value2 = arr19;

                //ObjWorkBookOut2.SaveAs();
                ObjWorkBookOut2.Close(null, null, null);
                toolStripProgressBar1.Value = 90;

                groupBox10.Visible = false;

                ObjExcel.Quit();

                this.uDislTableAdapter.Fill(this.uniq1DataSet.uDisl);
                toolStripProgressBar1.Value = 100;

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string bro_path = "C:\\Program Files\\Mozilla Firefox\\firefox.exe";
            string htm_path = "https://kuban.mts.ru/personal/sendsms";

            if (File.Exists("simple_sms_browser.txt"))
            {
                bro_path = File.ReadAllText("simple_sms_browser.txt");

            }
            if (File.Exists("simple_sms_https.txt"))
            {
                htm_path = File.ReadAllText("simple_sms_https.txt");

            }

            ProcessStartInfo web_browser = new ProcessStartInfo(bro_path, htm_path);
            Process.Start(web_browser);

        }
        private void загрузитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {

            ThreadStart thr2 = new ThreadStart(Load_fucking_deblo);
            Thread Load_deblo = new Thread(thr2);
            Load_deblo.Start();
        }

        private void реальноеВремяОхраныToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (реальноеВремяОхраныToolStripMenuItem.Checked)
            {
                groupBox11.Visible = false;
                реальноеВремяОхраныToolStripMenuItem.Checked = false;
            }
            else {
                groupBox11.Left = 4;
                groupBox11.Top = 658;
                groupBox11.Visible = true;
                реальноеВремяОхраныToolStripMenuItem.Checked = true;

            }

        }

        private void блокДанныхОбОбъектеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (блокДанныхОбОбъектеToolStripMenuItem.Checked)
            {
                groupBox3.Visible = false;
                блокДанныхОбОбъектеToolStripMenuItem.Checked = false;
            }
            else {
                groupBox3.Left = 4;
                groupBox3.Top = 380;
                groupBox3.Visible = true;
                блокДанныхОбОбъектеToolStripMenuItem.Checked = true;

            }

        }

        private void поискВСработкахToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (поискВСработкахToolStripMenuItem.Checked)
            {
                groupBox2.Visible = false;
                поискВСработкахToolStripMenuItem.Checked = false;
            }
            else {
                groupBox2.Left = 4;
                groupBox2.Top = 158;
                groupBox2.Visible = true;
                поискВСработкахToolStripMenuItem.Checked = true;

            }

        }

        private void поискВДислокацииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (поискВДислокацииToolStripMenuItem.Checked)
            {
                groupBox1.Visible = false;
                поискВДислокацииToolStripMenuItem.Checked = false;
            }
            else {
                groupBox1.Left = 4;
                groupBox1.Top = 12;
                groupBox1.Visible = true;
                поискВДислокацииToolStripMenuItem.Checked = true;

            }

        }

        private void времяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (времяToolStripMenuItem.Checked)
            {
                groupBox5.Visible = false;
                времяToolStripMenuItem.Checked = false;
            }
            else {
                groupBox5.Left = 5;
                groupBox5.Top = 903;
                groupBox5.Visible = true;
                времяToolStripMenuItem.Checked = true;
            }

        }

        private void погодныеУсловияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (погодныеУсловияToolStripMenuItem.Checked)
            {
                groupBox8.Visible = false;
                погодныеУсловияToolStripMenuItem.Checked = false;
            }
            else {
                groupBox8.Left = 4;
                groupBox8.Top = 1511;
                groupBox8.Visible = true;
                погодныеУсловияToolStripMenuItem.Checked = true;

            }
        }

        private void сработавшееОборкдованиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (сработавшееОборкдованиеToolStripMenuItem.Checked)
            {
                groupBox7.Visible = false;
                сработавшееОборкдованиеToolStripMenuItem.Checked = false;
            }
            else {
                groupBox7.Left = 4;
                groupBox7.Top = 1439;
                groupBox7.Visible = true;
                сработавшееОборкдованиеToolStripMenuItem.Checked = true;

            }
        }

        private void дебиторскаяToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (дебиторскаяToolStripMenuItem1.Checked)
            {
                дебиторскаяToolStripMenuItem.Visible = false;
                дебиторскаяToolStripMenuItem1.Checked = false;
            }
            else {
                дебиторскаяToolStripMenuItem.Visible = true;
                дебиторскаяToolStripMenuItem1.Checked = true;

            }

        }

        private void дислокацияToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (дислокацияToolStripMenuItem1.Checked)
            {
                дислокацияToolStripMenuItem1.Checked = false;
                дислокацияToolStripMenuItem.Visible = false;
            }
            else {
                дислокацияToolStripMenuItem.Visible = true;
                дислокацияToolStripMenuItem1.Checked = true;
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данные по фактическому времени охраны: \n \\\\192.168.51.1\\pco\\Овчаренко\\Переохрана\\ \n Только для систем Приток-А, Феникс (Сирень)", "pcn_b0x", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void добавитьАвтомобилиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            List<string> ini_disl_mob = new List<string>();
            object[,] arr1 = new object[5001, 1];
            object[,] arr2 = new object[5001, 1];
            object[,] arr3 = new object[5001, 1];
            object[,] arr4 = new object[5001, 1];
            object[,] arr5 = new object[5001, 1];
            object[,] arr6 = new object[5001, 1];
            object[,] arr7 = new object[5001, 1];
            object[,] arr8 = new object[5001, 1];
            object[,] arr9 = new object[5001, 1];
            object[,] arr10 = new object[5001, 1];
            object[,] arr11 = new object[5001, 1];


            if (File.Exists("path_disl_mob.txt"))
            {
                ini_disl_mob.AddRange((string[])File.ReadAllLines("path_disl_mob.txt", Encoding.GetEncoding(1251)));
            }

            if (!File.Exists(ini_disl_mob[0]))
            {
                MessageBox.Show("Нет файла: %автомобили% \n");
            }
            else {

                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(ini_disl_mob[0], 0, true, 5, "907", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range(ini_disl_mob[1] + "1", ini_disl_mob[1] + "5000");
                arr1 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[2] + "1", ini_disl_mob[2] + "5000");
                arr2 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[3] + "1", ini_disl_mob[3] + "5000");
                arr3 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[4] + "1", ini_disl_mob[4] + "5000");
                arr4 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[5] + "1", ini_disl_mob[5] + "5000");
                arr5 = (object[,])range.Value2;
                /*range = ObjWorkSheet.get_Range(ini_disl_mob[6] + "1", ini_disl_mob[6] + "5000");
                arr6 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[7] + "1", ini_disl_mob[7] + "5000");
                arr7 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[8] + "1", ini_disl_mob[8] + "5000");
                arr8 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[9] + "1", ini_disl_mob[9] + "5000");
                arr9 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[10] + "1", ini_disl_mob[10] + "5000");
                arr10 = (object[,])range.Value2;
                range = ObjWorkSheet.get_Range(ini_disl_mob[11] + "1", ini_disl_mob[11] + "5000");
                arr11 = (object[,])range.Value2;
                */
                ///////////////////////////////////////////////////ObjWorkBook.Close(null, null, null);

                for (int i = 1; i < 500; i++)
                {
                    if (arr2[i, 1] != null)
                    {
                        if (arr1[i, 1] == null) arr1[i, 1] = "(пустое поле)";
                        if (arr3[i, 1] == null) arr3[i, 1] = "(пустое поле)";
                        if (arr4[i, 1] == null) arr4[i, 1] = "(пустое поле)";
                        if (arr5[i, 1] == null) arr5[i, 1] = "(пустое поле)";

                        /*if (arr6[i, 1] == null) arr6[i, 1] = "(пусто)";
                        if (arr7[i, 1] == null) arr7[i, 1] = "(пусто)";
                        if (arr8[i, 1] == null) arr8[i, 1] = "-";
                        if (arr9[i, 1] == null) arr9[i, 1] = "-";
                        if (arr10[i, 1] == null) arr10[i, 1] = "нет";
                        if (arr11[i, 1] == null) arr11[i, 1] = "0,0";*/

                        if (arr1[i, 1].ToString().Contains("'")) arr1[i, 1] = arr1[i, 1].ToString().Replace('\'', ' ');
                        if (arr2[i, 1].ToString().Contains("'")) arr2[i, 1] = arr2[i, 1].ToString().Replace('\'', ' ');
                        if (arr3[i, 1].ToString().Contains("'")) arr3[i, 1] = arr3[i, 1].ToString().Replace('\'', ' ');
                        if (arr4[i, 1].ToString().Contains("'")) arr4[i, 1] = arr4[i, 1].ToString().Replace('\'', ' ');

                        /*if (arr6[i, 1].ToString().Contains("'")) arr6[i, 1] = arr6[i, 1].ToString().Replace('\'', ' ');
                        if (arr7[i, 1].ToString().Contains("'")) arr7[i, 1] = arr7[i, 1].ToString().Replace('\'', ' ');
                        if (arr8[i, 1].ToString().Contains("'")) arr8[i, 1] = arr8[i, 1].ToString().Replace('\'', ' ');
                        if (arr9[i, 1].ToString().Contains("'")) arr9[i, 1] = arr9[i, 1].ToString().Replace('\'', ' ');
                        if (arr10[i, 1].ToString().Contains("'")) arr10[i, 1] = arr10[i, 1].ToString().Replace('\'', ' ');
                        if (arr11[i, 1].ToString().Contains("'")) arr11[i, 1] = arr11[i, 1].ToString().Replace('\'', ' ');
                        */

                        SqlCommand comm2 = new SqlCommand();
                        comm2.Connection = Program.conn;
                        string sql_comm_str;

                        sql_comm_str = "INSERT [dbo].[uDisl] ([poz] ,[obj_org] ,[obj_name] ,[obj_type] ,[type_protect]) VALUES ('Алмаз " + arr1[i, 1].ToString() + "', '" + arr2[i, 1].ToString() + "', '" + arr3[i, 1].ToString() + " (" + arr5[i, 1].ToString() + ")" + "',1,2)";

                        comm2.CommandText = sql_comm_str;
                        comm2.ExecuteNonQuery();
                        //Dispose()
                    }
                }
                ObjWorkBook.Close();
                MessageBox.Show("Автомобили добавлены!");
                this.uDislTableAdapter.Fill(this.uniq1DataSet.uDisl);
            }
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text.Contains("аявка"))
            {
                //
                textBox13.Text = "(Заявка)";
                textBox14.Text = "00";
                comboBox11.Text = "(Заявка)";
                comboBox4.Text = "Заявка";
                comboBox7.Text = ".";
                comboBox8.Text = ".";

            }
        }

        private void выгрузитьВсеДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            this.uDepTableAdapter.Fill(this.uniq1DataSet.uDep);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // Технический вызов UPDATE запроса **************************************************
            // использовать однократно, отключить элемент "button9" ******************************
            // string sql_comm_str = "";
            //10. УТЗ
            //string xXx = "10. УТЗ"; // new
            //string yYy = "9.УТЗ";   // old

            //9. Другие
            //string xXx = "9. Другие"; // new
            //string yYy = "6. Другие";   // old

            //8.5 Вина. Другие
            //string xXx = "8.5 Вина. Другие"; // new
            //string yYy = "4.5. (Вина) Другие";   // old

            //8.4 Вина. Животные
            //string xXx = "8.4 Вина. Животные"; // new
            //string yYy = "4.4.(Вина)Животные, насек";   // old

            //8.3 Вина. Случайно ТС
            //string xXx = "8.3 Вина. Случайно ТС"; // new
            //string yYy = "4.3. (Вина) Случайно ТС";   // old

            //8.2 Вина. Неверн. действ.
            //string xXx = "8.2 Вина. Неверн. действ."; // new
            //string yYy = "4.2. (Вина) Неверн. дейст";   // old

            //8.1 Вина. Забыли снять
            //string xXx = "8.1 Вина. Забыли снять"; // new
            //string yYy = "4.1. (Вина) Забыли снять";   // old

            //7. Неудовлетв. ИТУ
            //string xXx = "7. Неудовлетв. ИТУ"; // new
            //string yYy = "Тех. укрепленность";   // old

            //6. Обестачивание
            //string xXx = "6. Обестачивание"; // new
            //string yYy = "3.Обестачивание";   // old

            //5. Вина АТС************************************************************
            //string xXx = "5. Вина АТС"; // new
            //string yYy = "2.Вина АТС";   // old

            //5. Вина АТС************************************************************
            //string xXx = "5. Вина АТС"; // new
            //string yYy = "2.Вина оператора связи";   // old

            //3. Неиспр. ППК
            //string xXx = "3. Неиспр. ППК"; // new
            //string yYy = "Неиспр. ППК";   // old

            //2. Неиспр. датчики
            //string xXx = "2. Неиспр. датчики"; // new
            //string yYy = "Неиспр. Датчики";   // old

            //1. Неиспр. МКИ и ШС
            //string xXx = "1. Неиспр. МКИ и ШС"; // new
            //string yYy = "Неиспр. Шлейфы, СМК";   // old

            /*SqlCommand imp_dep = new SqlCommand();
            imp_dep.Connection = Program.conn;
            sql_comm_str = "UPDATE [dbo].[uDep] SET [reason_start]='" + xXx + "' WHERE reason_start='" + yYy + "' "; 
            imp_dep.CommandText = sql_comm_str;
            imp_dep.ExecuteNonQuery();
            */
            /*
            SqlCommand imp_dep2 = new SqlCommand();
            imp_dep2.Connection = Program.conn;
            sql_comm_str = "UPDATE [dbo].[uDep] SET [reason_end]='" + xXx + "' WHERE reason_end='" + yYy + "' ";
            imp_dep2.CommandText = sql_comm_str;
            imp_dep2.ExecuteNonQuery();
            */
        }

        private void притокАToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object[,] arr14 = new object[10001, 1];
            object[,] arr15 = new object[10001, 1];
            object[,] arr16 = new object[10001, 1];
            object[,] arr17 = new object[10001, 1];
            object[,] arr18 = new object[10001, 1];
            object[,] arr19 = new object[10001, 1];

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.ShowDialog(this);

            //MessageBox.Show("Путь к файлу: " + fdlg.FileName);

            toolStripProgressBar1.Value = 5;

            if (File.Exists(fdlg.FileName) && fdlg.FileName.Contains("\\\\") == false)
            {
                int c = 0;
                SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                SqlCommand comm3 = new SqlCommand();
                comm3.Connection = Program.conn;

                string sql_comm_str = "";


                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut2 = ObjExcel.Workbooks.Open(fdlg.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOut2;
                ObjWorkSheetOut2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut2.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range3 = ObjWorkSheetOut2.get_Range("E" + "1", "E" + "10000");
                arr14 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("G" + "1", "G" + "10000");
                arr15 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("I" + "1", "I" + "10000");
                arr16 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("J" + "1", "J" + "10000");
                arr17 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("L" + "1", "L" + "10000");
                arr18 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("M" + "1", "M" + "10000");
                arr19 = (object[,])range3.Value2;

                for (int i = 1; i < 5000; i++)
                {
                    //progressBar1.Value++;
                    double progress1;
                    if (arr18[i, 1] == null) arr18[i, 1] = (object)" ";
                    if (arr19[i, 1] == null) arr19[i, 1] = (object)" ";

                    if (arr14[i, 1] != null && arr15[i, 1] != null /*&& arr14[i, 1].ToString() != "" && arr14[i, 1].ToString() != " "*/)
                    {
                        if (arr16[i, 1] == null) arr16[i, 1] = "(пустое поле)";
                        if (arr17[i, 1] == null) arr17[i, 1] = "(пустое поле)";

                        if (arr14[i, 1].ToString().Contains("'")) arr14[i, 1] = arr14[i, 1].ToString().Replace('\'', ' ');
                        if (arr15[i, 1].ToString().Contains("'")) arr15[i, 1] = arr15[i, 1].ToString().Replace('\'', ' ');
                        if (arr16[i, 1].ToString().Contains("'")) arr16[i, 1] = arr16[i, 1].ToString().Replace('\'', ' ');
                        if (arr17[i, 1].ToString().Contains("'")) arr17[i, 1] = arr17[i, 1].ToString().Replace('\'', ' ');
                        if (arr18[i, 1].ToString().Contains("'")) arr18[i, 1] = arr18[i, 1].ToString().Replace('\'', ' ');
                        if (arr19[i, 1].ToString().Contains("'")) arr19[i, 1] = arr19[i, 1].ToString().Replace('\'', ' ');

                        sql_comm_str = @"SELECT * FROM  [dbo].[uPritokA] WHERE  [poz]='" + arr14[i, 1].ToString() + "' AND [date1]='" + arr15[i, 1].ToString() + "'";

                        comm2.CommandText = sql_comm_str;
                        SqlDataReader test_r1 = comm2.ExecuteReader();

                        if (test_r1.Read() == false)
                        {
                            test_r1.Close();
                            sql_comm_str = "INSERT [dbo].[uPritokA] ([poz] ,[date1] ,[obj_name] ,[obj_addr] ,[temp1] ,[temp2]) VALUES ('" + arr14[i, 1].ToString() + "', '" + arr15[i, 1].ToString() + "', '" + arr16[i, 1].ToString() + "', '" + arr17[i, 1].ToString() + "', '" + arr18[i, 1].ToString() + "', '" + arr19[i, 1].ToString() + "')";

                            comm3.CommandText = sql_comm_str;
                            comm3.ExecuteNonQuery();
                            c++; //oooohhhh yeeaaaaaaaa 
                        }
                        else test_r1.Close();
                    }
                }
                toolStripProgressBar1.Value = 100;
                ObjWorkBookOut2.Close(null, null, null);
                MessageBox.Show("Добавлено тревог: " + c.ToString(), "Приток-А, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
        private void загрузитьТревогиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
        }

        private void притокАToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            List<string> pri1 = new List<string>();
            if (File.Exists("pri1.txt"))
            {
                pri1.AddRange((string[])File.ReadAllLines("pri1.txt", Encoding.GetEncoding(1251)));
            }

            List<string> pri2 = new List<string>();
            if (File.Exists("pri2.txt"))
            {
                pri2.AddRange((string[])File.ReadAllLines("pri2.txt", Encoding.GetEncoding(1251)));
            }

            List<string> pri3 = new List<string>();
            if (File.Exists("pri3.txt"))
            {
                pri3.AddRange((string[])File.ReadAllLines("pri3.txt", Encoding.GetEncoding(1251)));
            }

            List<string> ini_dates_yupi8 = new List<string>();
            if (File.Exists("pritonA.dates"))
            {
                ini_dates_yupi8.AddRange((string[])File.ReadAllLines("pritonA.dates", Encoding.GetEncoding(1251)));
            }
            //woEthernet
            List<string> ini_woE_yupi8 = new List<string>();
            if (File.Exists("pritonA.woEthernet"))
            {
                ini_woE_yupi8.AddRange((string[])File.ReadAllLines("pritonA.woEthernet", Encoding.GetEncoding(1251)));
            }

            //
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;
            string sql_comm_str;

            sql_comm_str = @"SELECT * FROM  [dbo].[uPritokA] WHERE [id]>1 ORDER BY [id]";

            comm2.CommandText = sql_comm_str;
            SqlDataReader test_r1 = comm2.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;
            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];

                rowcount++;
            }
            test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();

            //Создаём приложение.

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut = ObjExcel.Workbooks.Add(xls_path + "PritokA1.xls");

            //progressBar1.Value = 25;
            //label1.Text = "Run Excel App...";
            Refresh();

            object[,] arr1 = new object[rowcount, 29];
            object[,] arr3 = new object[rowcount, 29];
            int rowcount1 = rowcount;


            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[1];

            if (ws == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {
                for (int i = 0; i < rowcount; i++)
                {
                    char[] ch = new char[2];
                    // = {' ','.'};
                    ch[0] = ' ';
                    ch[1] = '.';
                    string[] strm = new string[10];
                    strm = ds1[i, 2].ToString().Split(ch);


                    arr1[i, 0] = ds1[i, 3].ToString();
                    arr1[i, 1] = ds1[i, 4].ToString();
                    if (strm.Length > 2)
                        arr1[i, 2] = strm[2] + "." + strm[1] + "." + strm[0];//ds1[i, 2].ToString();
                    else
                        arr1[i, 2] = "неверные данные";
                    arr1[i, 3] = ds1[i, 1].ToString();
                    arr1[i, 4] = ds1[i, 0].ToString();
                    
                    arr1[i, 5] = ds1[i, 5].ToString();
                    arr1[i, 6] = ds1[i, 6].ToString();

                }

                int k = rowcount + 2;
                ws.get_Range("A" + 3, "AA" + k).Value2 = arr1;

            }

            /////////////////////////////////////////******************************************************/////////////////////////////////////////////

            sql_comm_str = @"SELECT [poz],[temp1]  ,COUNT([poz]) as c_poz
                            FROM [dbo].[uPritokA]
                            WHERE [id]>1
                             GROUP BY [poz],[temp1] 
                            HAVING COUNT([poz]) > 1 " +
                " ORDER BY c_poz";

            comm2.CommandText = sql_comm_str;
            test_r1 = comm2.ExecuteReader();


            Refresh();

            object[,] ds2 = new object[150001, 4];
            rowcount = 0;
            while (test_r1.Read())
            {
                ds2[rowcount, 0] = test_r1[0];
                ds2[rowcount, 1] = test_r1[1];
                ds2[rowcount, 2] = test_r1[2];

                rowcount++;
            }
            test_r1.Close();

            object[,] arr2 = new object[rowcount, 72];

            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[3];

            if (ws2 == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {

                for (int i = 0; i < rowcount; i++)
                {
                    /*if (!checkBox1.Checked && ds2[i, 1].ToString() == "Заявка")
                    {

                        continue;
                    }*/
                    for (int j = 0; j < rowcount1; j++)
                    {
                        if (ds2[i, 0].ToString() == arr1[j, 3].ToString())
                        {
                            arr2[i, 0] = arr1[j, 0];
                            arr2[i, 1] = arr1[j, 1];
                            //[serv_org], [dogovor], [tel_info], [deblo_otvetish]
                            sql_comm_str = @"SELECT [serv_org],[deblo_otvetish],[tel_info] 
                            FROM [dbo].[uDisl]
                            WHERE [poz] LIKE '" + "П " + arr1[j, 3].ToString() + "'";

                            comm2.CommandText = sql_comm_str;
                            test_r1 = comm2.ExecuteReader();
                            if (test_r1.Read())
                            {

                                arr2[i, 5] = test_r1[0].ToString();
                                arr2[i, 6] = test_r1[1].ToString();
                                arr2[i, 7] = test_r1[2].ToString();

                                test_r1.Close();
                            }
                            else test_r1.Close();

                            int l = 0;
                            int p = 0;
                            for (l = 0; l < ini_dates_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if ( arr1[j, 5].ToString()=="Тревога АН - авария направления" && arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 2].ToString() == ini_dates_yupi8[l]/*arr1[j, 2].ToString() == arr1[p, 2].ToString()*/)
                                    {
                                        arr2[i, 8 + l] = "н";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < ini_woE_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString() == ini_woE_yupi8[l]/*arr1[j, 2].ToString() == arr1[p, 2].ToString()*/)
                                    {
                                        arr2[i, 70] = "нет проводного Интернета!";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < pri1.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString() == pri1[l])
                                    {
                                        arr2[i, 71] = "111";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < pri2.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString() == pri2[l])
                                    {
                                        arr2[i, 71] = "222";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < pri3.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString() == pri3[l])
                                    {
                                        arr2[i, 71] = "333";
                                        break;
                                    }

                                }

                            }

                            break;
                        }
                    }

                    arr2[i, 2] = ds2[i, 0].ToString();
                    arr2[i, 4] = ds2[i, 1].ToString();
                    arr2[i, 3] = ds2[i, 2].ToString();

                }

                int m = 0;
                for (int n = 0; n < rowcount; n++)
                {
                    if (arr2[n, 4] != null)
                    {
                        arr2[m, 0] = arr2[n, 0];
                        arr2[m, 1] = arr2[n, 1];
                        arr2[m, 2] = arr2[n, 2];
                        arr2[m, 3] = arr2[n, 3];
                        arr2[m, 4] = arr2[n, 4];
                        arr2[m, 5] = arr2[n, 5];
                        arr2[m, 6] = arr2[n, 6];
                        arr2[m, 7] = arr2[n, 7];
                        m++;
                    }

                }
                int k = m + 2;
                ws2.get_Range("A" + 3, "BT" + k).Value2 = arr2;

                object[,] arr4 = new object[1, 71];
                for (int o = 0; o < ini_dates_yupi8.Count(); o++)
                {
                    arr4[0, o] = (object)ini_dates_yupi8[o];
                }
                ws2.get_Range("I" + 1, "BT" + 1).Value2 = arr4;

            }


            //progressBar1.Value = 100;
            //label1.Text = "Open excel window. Bye bye... ";
            Refresh();


            ObjExcel.Visible = true;

        }

        private void изСпискаПритокАToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;

            string sql_comm_str = "";

            sql_comm_str = "TRUNCATE TABLE [dbo].[uPritokA]";

            comm2.CommandText = sql_comm_str;
            comm2.ExecuteNonQuery();
            MessageBox.Show("Список тревог очищен полностью!!!", "Приток-А, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void отчетыПоТревогамToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void юпитер8ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            object[,] arr14 = new object[10001, 1];
            object[,] arr15 = new object[10001, 1];
            object[,] arr16 = new object[10001, 1];
            object[,] arr17 = new object[10001, 1];
            object[,] arr18 = new object[10001, 1];
            object[,] arr19 = new object[10001, 1];

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.ShowDialog(this);

            //MessageBox.Show("Путь к файлу: " + fdlg.FileName);

            toolStripProgressBar1.Value = 5;

            if (File.Exists(fdlg.FileName) && fdlg.FileName.Contains("\\\\") == false)
            {
                int c = 0;
                SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                SqlCommand comm3 = new SqlCommand();
                comm3.Connection = Program.conn;

                string sql_comm_str = "";


                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut2 = ObjExcel.Workbooks.Open(fdlg.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOut2;
                ObjWorkSheetOut2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut2.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range3 = ObjWorkSheetOut2.get_Range("B" + "1", "B" + "10000");
                arr14 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("A" + "1", "A" + "10000");
                arr15 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("C" + "1", "C" + "10000");
                arr16 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("H" + "1", "H" + "10000");
                arr17 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("F" + "1", "F" + "10000");
                arr18 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("E" + "1", "E" + "10000");
                arr19 = (object[,])range3.Value2;

                char[] ch = new char[2];
                // = {' ','.'};
                ch[0] = ' ';
                ch[1] = '.';
                string[] strm = new string[10];
                strm = fdlg.FileName.Split(ch);

                for (int i = 1; i < 5000; i++)
                {
                    //progressBar1.Value++;
                    double progress1;

                    arr15[i, 1] = (object)(strm[3] + "." + strm[2] + "." + strm[1]);


                    if (arr18[i, 1] == null) arr18[i, 1] = (object)" ";
                    if (arr19[i, 1] == null) arr19[i, 1] = (object)" ";

                    if (arr14[i, 1] != null && arr15[i, 1] != null /*&& arr14[i, 1].ToString() != "" && arr14[i, 1].ToString() != " "*/)
                    {
                        if (arr16[i, 1] == null) arr16[i, 1] = "(пустое поле)";
                        if (arr17[i, 1] == null) arr17[i, 1] = "(пустое поле)";

                        if (arr14[i, 1].ToString().Contains("'")) arr14[i, 1] = arr14[i, 1].ToString().Replace('\'', ' ');
                        if (arr15[i, 1].ToString().Contains("'")) arr15[i, 1] = arr15[i, 1].ToString().Replace('\'', ' ');
                        if (arr16[i, 1].ToString().Contains("'")) arr16[i, 1] = arr16[i, 1].ToString().Replace('\'', ' ');
                        if (arr17[i, 1].ToString().Contains("'")) arr17[i, 1] = arr17[i, 1].ToString().Replace('\'', ' ');
                        if (arr18[i, 1].ToString().Contains("'")) arr18[i, 1] = arr18[i, 1].ToString().Replace('\'', ' ');
                        if (arr19[i, 1].ToString().Contains("'")) arr19[i, 1] = arr19[i, 1].ToString().Replace('\'', ' ');

                        sql_comm_str = @"SELECT * FROM  [dbo].[uJupiter8] WHERE  [poz]='" + arr14[i, 1].ToString() + "' AND [date1]='" + arr15[i, 1].ToString() + "'";

                        comm2.CommandText = sql_comm_str;
                        SqlDataReader test_r1 = comm2.ExecuteReader();

                        if (test_r1.Read() == false)
                        {
                            test_r1.Close();
                            sql_comm_str = "INSERT [dbo].[uJupiter8] ([poz] ,[date1] ,[obj_name] ,[obj_addr] ,[temp1] ,[temp2]) VALUES ('" + arr14[i, 1].ToString() + "', '" + arr15[i, 1].ToString() + "', '" + arr16[i, 1].ToString() + "', '" + arr17[i, 1].ToString() + "', '" + arr18[i, 1].ToString() + "', '" + arr19[i, 1].ToString() + "')";

                            comm3.CommandText = sql_comm_str;
                            comm3.ExecuteNonQuery();
                            c++; //oooohhhh yeeaaaaaaaa 
                        }
                        else test_r1.Close();
                    }
                }
                toolStripProgressBar1.Value = 100;
                ObjWorkBookOut2.Close(null, null, null);
                MessageBox.Show("Добавлено тревог: " + c.ToString(), "Юпитер 9, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

        }

        private void изСпискаЮпитер8ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;

            string sql_comm_str = "";

            sql_comm_str = "TRUNCATE TABLE [dbo].[uJupiter8]";

            comm2.CommandText = sql_comm_str;
            comm2.ExecuteNonQuery();
            MessageBox.Show("Список тревог очищен полностью!!!", "Юпитер 9, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        private void юпитер8ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            List<string> pri1 = new List<string>();
            if (File.Exists("yupi1.txt"))
            {
                pri1.AddRange((string[])File.ReadAllLines("yupi1.txt", Encoding.GetEncoding(1251)));
            }

            List<string> pri2 = new List<string>();
            if (File.Exists("yupi2.txt"))
            {
                pri2.AddRange((string[])File.ReadAllLines("yupi2.txt", Encoding.GetEncoding(1251)));
            }

            List<string> pri3 = new List<string>();
            if (File.Exists("yupi3.txt"))
            {
                pri3.AddRange((string[])File.ReadAllLines("yupi3.txt", Encoding.GetEncoding(1251)));
            }

            //
            List<string> ini_dates_yupi8 = new List<string>();
            if (File.Exists("yupi8.dates"))
            {
                ini_dates_yupi8.AddRange((string[])File.ReadAllLines("yupi8.dates", Encoding.GetEncoding(1251)));
            }

            //
            List<string> ini_woE_yupi8 = new List<string>();
            if (File.Exists("yupi8.woEthernet"))
            {
                ini_woE_yupi8.AddRange((string[])File.ReadAllLines("yupi8.woEthernet", Encoding.GetEncoding(1251)));
            }


            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;
            string sql_comm_str;

            sql_comm_str = @"SELECT * FROM  [dbo].[uJupiter8] WHERE [id]>1 ORDER BY [id]";

            comm2.CommandText = sql_comm_str;
            SqlDataReader test_r1 = comm2.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;
            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];

                rowcount++;
            }
            test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();

            //Создаём приложение.

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut = ObjExcel.Workbooks.Add(xls_path + "PritokA1.xls");

            //progressBar1.Value = 25;
            //label1.Text = "Run Excel App...";
            Refresh();

            object[,] arr1 = new object[rowcount, 29];
            object[,] arr3 = new object[rowcount, 29];
            int rowcount1 = rowcount;


            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[1];

            if (ws == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {
                for (int i = 0; i < rowcount; i++)
                {
                    arr1[i, 0] = ds1[i, 3].ToString();
                    arr1[i, 1] = ds1[i, 4].ToString();
                    arr1[i, 2] = ds1[i, 2].ToString();
                    arr1[i, 3] = ds1[i, 1].ToString();
                    arr1[i, 4] = ds1[i, 0].ToString();
                    arr1[i, 5] = ds1[i, 5].ToString();
                    arr1[i, 6] = ds1[i, 6].ToString();

                }

                int k = rowcount + 2;
                ws.get_Range("A" + 3, "AA" + k).Value2 = arr1;

            }

            /////////////////////////////////////////******************************************************/////////////////////////////////////////////

            sql_comm_str = @"SELECT [poz],[temp1]  ,COUNT([poz]) as c_poz
                            FROM [dbo].[uJupiter8]
                            WHERE [id]>1
                             GROUP BY [poz],[temp1] 
                            HAVING COUNT([poz]) > 1 " +
                " ORDER BY c_poz";

            comm2.CommandText = sql_comm_str;
            test_r1 = comm2.ExecuteReader();


            Refresh();

            object[,] ds2 = new object[150001, 4];
            int rowcount2 = rowcount;
            rowcount = 0;
            while (test_r1.Read())
            {
                ds2[rowcount, 0] = test_r1[0];
                ds2[rowcount, 1] = test_r1[1];
                ds2[rowcount, 2] = test_r1[2];

                rowcount++;
            }
            test_r1.Close();

            object[,] arr2 = new object[rowcount, 72];

            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[3];

            if (ws2 == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {

                for (int i = 0; i < rowcount; i++)
                {
                    /*if (!checkBox1.Checked && ds2[i, 1].ToString() == "Заявка")
                    {

                        continue;
                    }*/
                    for (int j = 0; j < rowcount1; j++)
                    {
                        if (ds2[i, 0].ToString() == arr1[j, 3].ToString())
                        {
                            arr2[i, 0] = arr1[j, 0];
                            arr2[i, 1] = arr1[j, 1];

                                //[serv_org], [dogovor], [tel_info], [deblo_otvetish]
                                sql_comm_str = @"SELECT [serv_org],[deblo_otvetish],[tel_info] 
                            FROM [dbo].[uDisl]
                            WHERE [poz] LIKE '" + "Ю " + arr1[j, 3].ToString() + "'";

                            comm2.CommandText = sql_comm_str;
                            test_r1 = comm2.ExecuteReader();
                            if (test_r1.Read())
                            {

                                arr2[i, 5] = test_r1[0].ToString();
                                arr2[i, 6] = test_r1[1].ToString();
                                arr2[i, 7] = test_r1[2].ToString();

                                test_r1.Close();
                            }
                            else test_r1.Close();

                            int l = 0;
                            int p = 0;
                            for (l = 0; l < ini_dates_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount2; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 2].ToString() == ini_dates_yupi8[l]/*arr1[j, 2].ToString() == arr1[p, 2].ToString()*/)
                                    {
                                        arr2[i, 8 + l] = "н";
                                        break;
                                    }

                                }

                            }
                            char[] ch = new char[1];
                            ch[0] = '0';

                            for (l = 0; l < ini_woE_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString().TrimStart(ch) == ini_woE_yupi8[l]/*arr1[j, 2].ToString() == arr1[p, 2].ToString()*/)
                                    {
                                        arr2[i, 70] = "нет проводного Интернета!";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < pri1.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString().TrimStart(ch) == pri1[l])
                                    {
                                        arr2[i, 71] = "111";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < pri2.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString().TrimStart(ch) == pri2[l])
                                    {
                                        arr2[i, 71] = "222";
                                        break;
                                    }

                                }

                            }
                            for (l = 0; l < pri3.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString().TrimStart(ch) == pri3[l])
                                    {
                                        arr2[i, 71] = "333";
                                        break;
                                    }

                                }

                            }

                            break;
                        }
                    }

                    arr2[i, 2] = ds2[i, 0].ToString();
                    arr2[i, 4] = ds2[i, 1].ToString();
                    arr2[i, 3] = ds2[i, 2].ToString();

                }

                int m = 0;
                for (int n = 0; n < rowcount; n++)
                {
                    if (arr2[n, 4] != null)
                    {
                        arr2[m, 0] = arr2[n, 0];
                        arr2[m, 1] = arr2[n, 1];
                        arr2[m, 2] = arr2[n, 2];
                        arr2[m, 3] = arr2[n, 3];
                        arr2[m, 4] = arr2[n, 4];
                        arr2[m, 5] = arr2[n, 5];
                        arr2[m, 6] = arr2[n, 6];
                        arr2[m, 7] = arr2[n, 7];
                        m++;
                    }

                }
                int k = m + 2;
                ws2.get_Range("A" + 3, "BT" + k).Value2 = arr2;
                object[,] arr4 = new object[1, 71];
                for (int o = 0; o < ini_dates_yupi8.Count(); o++)
                {
                    arr4[0, o] = (object)ini_dates_yupi8[o];
                }
                ws2.get_Range("I"+ 1, "BT"+ 1).Value2 = arr4;

                //progressBar1.Value = 100;
                //label1.Text = "Open excel window. Bye bye... ";
                Refresh();

            }
            ObjExcel.Visible = true;

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (textBox23.Text == "Ф")
            {
                string sql_comm_str = "";
                SqlCommand imp_dep = new SqlCommand();
                imp_dep.Connection = Program.conn;

                sql_comm_str = "DELETE FROM [dbo].[uPhrenix] WHERE id between " + textBox21.Text + " and " + textBox22.Text;

                imp_dep.CommandText = sql_comm_str;
                imp_dep.ExecuteNonQuery();
            }
            if (textBox23.Text == "Ю")
            {
                string sql_comm_str = "";
                SqlCommand imp_dep = new SqlCommand();
                imp_dep.Connection = Program.conn;

                sql_comm_str = "DELETE FROM [dbo].[uJupiter8] WHERE id between " + textBox21.Text + " and " + textBox22.Text;

                imp_dep.CommandText = sql_comm_str;
                imp_dep.ExecuteNonQuery();
            }
            if (textBox23.Text == "П")
            {
                string sql_comm_str = "";
                SqlCommand imp_dep = new SqlCommand();
                imp_dep.Connection = Program.conn;

                sql_comm_str = "DELETE FROM [dbo].[uPritokA] WHERE id between " + textBox21.Text + " and " + textBox22.Text;

                imp_dep.CommandText = sql_comm_str;
                imp_dep.ExecuteNonQuery();
            }
        }

        private void интегралToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            object[,] arr14 = new object[10001, 1];
            object[,] arr15 = new object[10001, 1];
            object[,] arr16 = new object[10001, 1];
            object[,] arr17 = new object[10001, 1];
            object[,] arr18 = new object[10001, 1];
            object[,] arr19 = new object[10001, 1];

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.ShowDialog(this);

            //MessageBox.Show("Путь к файлу: " + fdlg.FileName);

            toolStripProgressBar1.Value = 5;

            if (File.Exists(fdlg.FileName) && fdlg.FileName.Contains("\\\\") == false)
            {
                int c = 0;
                SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                SqlCommand comm3 = new SqlCommand();
                comm3.Connection = Program.conn;

                string sql_comm_str = "";


                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut2 = ObjExcel.Workbooks.Open(fdlg.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOut2;
                ObjWorkSheetOut2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut2.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range3 = ObjWorkSheetOut2.get_Range("A" + "1", "A" + "10000");
                arr14 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("F" + "1", "F" + "10000");
                arr15 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("B" + "1", "B" + "10000");
                arr16 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("B" + "1", "B" + "10000");
                arr17 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("C" + "1", "C" + "10000");
                arr18 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("C" + "1", "C" + "10000");
                arr19 = (object[,])range3.Value2;

                char[] ch = new char[2];
                // = {' ','.'};
                ch[0] = ' ';
                ch[1] = '.';
                string[] strm = new string[10];
                strm = fdlg.FileName.Split(ch);

                for (int i = 1; i < 5000; i++)
                {
                    //progressBar1.Value++;
                    double progress1;

                    arr15[i, 1] = (object)(strm[3] + "." + strm[2] + "." + strm[1]);


                    if (arr18[i, 1] == null) arr18[i, 1] = (object)" ";
                    if (arr19[i, 1] == null) arr19[i, 1] = (object)" ";

                    if (arr14[i, 1] != null && arr15[i, 1] != null /*&& arr14[i, 1].ToString() != "" && arr14[i, 1].ToString() != " "*/)
                    {
                        string[] strch = new string[2];
                        // = {' ','.'};
                        strch[0] = "адрес:";
                        strch[1] = "IP";
                        string[] strm2 = new string[10];


                        if (arr16[i, 1] == null) arr16[i, 1] = "(пустое поле)";
                        if (arr17[i, 1] == null) arr17[i, 1] = "(пустое поле)";

                        strm2 = arr16[i, 1].ToString().Split(strch, StringSplitOptions.None);

                        if (strm2.Length>1)
                        {
                            arr16[i, 1] = strm2[0];
                            arr17[i, 1] = strm2[1];
                        }

                        strm2 = arr18[i, 1].ToString().Split(strch, StringSplitOptions.None);

                        if (strm2.Length > 1)
                        {
                            arr18[i, 1] = strm2[0];
                            arr19[i, 1] = strm2[1];
                        }



                        if (arr14[i, 1].ToString().Contains("'")) arr14[i, 1] = arr14[i, 1].ToString().Replace('\'', ' ');
                        if (arr15[i, 1].ToString().Contains("'")) arr15[i, 1] = arr15[i, 1].ToString().Replace('\'', ' ');
                        if (arr16[i, 1].ToString().Contains("'")) arr16[i, 1] = arr16[i, 1].ToString().Replace('\'', ' ');
                        if (arr17[i, 1].ToString().Contains("'")) arr17[i, 1] = arr17[i, 1].ToString().Replace('\'', ' ');
                        if (arr18[i, 1].ToString().Contains("'")) arr18[i, 1] = arr18[i, 1].ToString().Replace('\'', ' ');
                        if (arr19[i, 1].ToString().Contains("'")) arr19[i, 1] = arr19[i, 1].ToString().Replace('\'', ' ');

                        sql_comm_str = @"SELECT * FROM  [dbo].[uIntegralEvp] WHERE  [poz]='" + arr14[i, 1].ToString() + "' AND [date1]='" + arr15[i, 1].ToString() + "'";

                        comm2.CommandText = sql_comm_str;
                        SqlDataReader test_r1 = comm2.ExecuteReader();

                        if (test_r1.Read() == false)
                        {
                            test_r1.Close();
                            sql_comm_str = "INSERT [dbo].[uIntegralEvp] ([poz] ,[date1] ,[obj_name] ,[obj_addr] ,[temp1] ,[temp2]) VALUES ('" + arr14[i, 1].ToString() + "', '" + arr15[i, 1].ToString() + "', '" + arr16[i, 1].ToString() + "', '" + arr17[i, 1].ToString() + "', '" + arr18[i, 1].ToString() + "', '" + arr19[i, 1].ToString() + "')";

                            comm3.CommandText = sql_comm_str;
                            comm3.ExecuteNonQuery();
                            c++; //oooohhhh yeeaaaaaaaa 
                        }
                        else test_r1.Close();
                    }
                }
                toolStripProgressBar1.Value = 100;
                ObjWorkBookOut2.Close(null, null, null);
                MessageBox.Show("Добавлено тревог: " + c.ToString(), "Интеграл Евпатория, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        private void изСпискаИнтегралToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;

            string sql_comm_str = "";

            sql_comm_str = "TRUNCATE TABLE [dbo].[uIntegralEvp]";

            comm2.CommandText = sql_comm_str;
            comm2.ExecuteNonQuery();
            MessageBox.Show("Список тревог очищен полностью!!!", "Интеграл Евпатория, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        private void интегралToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            List<string> ini_dates_yupi8 = new List<string>();
            if (File.Exists("intEvp.dates"))
            {
                ini_dates_yupi8.AddRange((string[])File.ReadAllLines("intEvp.dates", Encoding.GetEncoding(1251)));
            }

            //
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;
            string sql_comm_str;

            sql_comm_str = @"SELECT * FROM  [dbo].[uIntegralEvp] WHERE [id]>1 ORDER BY [id]";

            comm2.CommandText = sql_comm_str;
            SqlDataReader test_r1 = comm2.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;
            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];

                rowcount++;
            }
            test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();

            //Создаём приложение.

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut = ObjExcel.Workbooks.Add(xls_path + "PritokA1.xls");

            //progressBar1.Value = 25;
            //label1.Text = "Run Excel App...";
            Refresh();

            object[,] arr1 = new object[rowcount, 29];
            object[,] arr3 = new object[rowcount, 29];
            int rowcount1 = rowcount;


            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[1];

            if (ws == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {
                for (int i = 0; i < rowcount; i++)
                {
                    arr1[i, 0] = ds1[i, 3].ToString();
                    arr1[i, 1] = ds1[i, 4].ToString();
                    arr1[i, 2] = ds1[i, 2].ToString();
                    arr1[i, 3] = ds1[i, 1].ToString();
                    arr1[i, 4] = ds1[i, 0].ToString();
                    arr1[i, 5] = ds1[i, 5].ToString();
                    arr1[i, 6] = ds1[i, 6].ToString();

                }

                int k = rowcount + 2;
                ws.get_Range("A" + 3, "AA" + k).Value2 = arr1;

            }

            /////////////////////////////////////////******************************************************/////////////////////////////////////////////

            sql_comm_str = @"SELECT [poz],[temp1]  ,COUNT([poz]) as c_poz
                            FROM [dbo].[uIntegralEvp]
                            WHERE [id]>1
                             GROUP BY [poz],[temp1] 
                            HAVING COUNT([poz]) > 1 " +
                " ORDER BY c_poz";

            comm2.CommandText = sql_comm_str;
            test_r1 = comm2.ExecuteReader();


            Refresh();

            object[,] ds2 = new object[150001, 4];
            rowcount = 0;
            while (test_r1.Read())
            {
                ds2[rowcount, 0] = test_r1[0];
                ds2[rowcount, 1] = test_r1[1];
                ds2[rowcount, 2] = test_r1[2];

                rowcount++;
            }
            test_r1.Close();

            object[,] arr2 = new object[rowcount, 40];

            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[3];

            if (ws2 == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {

                for (int i = 0; i < rowcount; i++)
                {
                    /*if (!checkBox1.Checked && ds2[i, 1].ToString() == "Заявка")
                    {

                        continue;
                    }*/
                    for (int j = 0; j < rowcount1; j++)
                    {
                        if (ds2[i, 0].ToString() == arr1[j, 3].ToString())
                        {
                            arr2[i, 0] = arr1[j, 0];
                            arr2[i, 1] = arr1[j, 1];
                            //[serv_org], [dogovor], [tel_info], [deblo_otvetish]
                            sql_comm_str = @"SELECT [serv_org],[deblo_otvetish],[tel_info] 
                            FROM [dbo].[uDisl]
                            WHERE [poz] LIKE '" + "А " + arr1[j, 3].ToString() + "'";

                            comm2.CommandText = sql_comm_str;
                            test_r1 = comm2.ExecuteReader();
                            if (test_r1.Read())
                            {

                                arr2[i, 5] = test_r1[0].ToString();
                                arr2[i, 6] = test_r1[1].ToString();
                                arr2[i, 7] = test_r1[2].ToString();

                                test_r1.Close();
                            }
                            else test_r1.Close();

                            int l = 0;
                            int p = 0;
                            for (l = 0; l < ini_dates_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 2].ToString() == ini_dates_yupi8[l]/*arr1[j, 2].ToString() == arr1[p, 2].ToString()*/)
                                    {
                                        arr2[i, 8 + l] = "н";
                                        break;
                                    }

                                }

                            }
                            /*for (l = 0; l < ini_woE_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 2].ToString() == ini_dates_yupi8[l])
                                    {
                                        arr2[i, 8 + l] = "н";
                                        break;
                                    }

                                }

                            }*/

                            break;
                        }
                    }

                    arr2[i, 2] = ds2[i, 0].ToString();
                    arr2[i, 4] = ds2[i, 1].ToString();
                    arr2[i, 3] = ds2[i, 2].ToString();

                }

                int m = 0;
                for (int n = 0; n < rowcount; n++)
                {
                    if (arr2[n, 4] != null)
                    {
                        arr2[m, 0] = arr2[n, 0];
                        arr2[m, 1] = arr2[n, 1];
                        arr2[m, 2] = arr2[n, 2];
                        arr2[m, 3] = arr2[n, 3];
                        arr2[m, 4] = arr2[n, 4];
                        arr2[m, 5] = arr2[n, 5];
                        arr2[m, 6] = arr2[n, 6];
                        arr2[m, 7] = arr2[n, 7];
                        m++;
                    }

                }
                int k = m + 2;
                ws2.get_Range("A" + 3, "AN" + k).Value2 = arr2;
                object[,] arr4 = new object[1, 40];
                for (int o = 0; o < ini_dates_yupi8.Count(); o++)
                {
                    arr4[0, o] = (object)ini_dates_yupi8[o];
                }
                ws2.get_Range("I" + 1, "AN" + 1).Value2 = arr4;

            }


            //progressBar1.Value = 100;
            //label1.Text = "Open excel window. Bye bye... ";
            Refresh();


            ObjExcel.Visible = true;

        }

        private void изСпискаФрениксToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;

            string sql_comm_str = "";

            sql_comm_str = "TRUNCATE TABLE [dbo].[uPhrenix]";

            comm2.CommandText = sql_comm_str;
            comm2.ExecuteNonQuery();
            MessageBox.Show("Список тревог очищен полностью!!!", "ФРЕНИКС!, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        private void френиксToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
            object[,] arr14 = new object[10001, 1];
            object[,] arr15 = new object[10001, 1];
            object[,] arr16 = new object[10001, 1];
            object[,] arr17 = new object[10001, 1];
            object[,] arr18 = new object[10001, 1];
            object[,] arr19 = new object[10001, 1];

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.ShowDialog(this);

            toolStripProgressBar1.Value = 5;

            if (File.Exists(fdlg.FileName) && fdlg.FileName.Contains("\\\\") == false)
            {
                int c = 0;
                SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                SqlCommand comm3 = new SqlCommand();
                comm3.Connection = Program.conn;

                string sql_comm_str = "";


                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut2 = ObjExcel.Workbooks.Open(fdlg.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOut2;
                ObjWorkSheetOut2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut2.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range3 = ObjWorkSheetOut2.get_Range("B" + "1", "B" + "10000");
                arr14 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("E" + "1", "E" + "10000");
                arr15 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("C" + "1", "C" + "10000");
                arr16 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("D" + "1", "D" + "10000");
                arr17 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("K" + "1", "K" + "10000");
                arr18 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("F" + "1", "F" + "10000");
                arr19 = (object[,])range3.Value2;

                for (int i = 1; i < 5000; i++)
                {
                    //progressBar1.Value++;
                    double progress1;
                    if (arr18[i, 1] == null) arr18[i, 1] = (object)" ";
                    if (arr19[i, 1] == null) arr19[i, 1] = (object)" ";

                    if (arr14[i, 1] != null && arr15[i, 1] != null /*&& arr14[i, 1].ToString() != "" && arr14[i, 1].ToString() != " "*/)
                    {
                        if (arr16[i, 1] == null) arr16[i, 1] = "(пустое поле)";
                        if (arr17[i, 1] == null) arr17[i, 1] = "(пустое поле)";

                        if (arr14[i, 1].ToString().Contains("'")) arr14[i, 1] = arr14[i, 1].ToString().Replace('\'', ' ');
                        if (arr15[i, 1].ToString().Contains("'")) arr15[i, 1] = arr15[i, 1].ToString().Replace('\'', ' ');
                        if (arr16[i, 1].ToString().Contains("'")) arr16[i, 1] = arr16[i, 1].ToString().Replace('\'', ' ');
                        if (arr17[i, 1].ToString().Contains("'")) arr17[i, 1] = arr17[i, 1].ToString().Replace('\'', ' ');
                        if (arr18[i, 1].ToString().Contains("'")) arr18[i, 1] = arr18[i, 1].ToString().Replace('\'', ' ');
                        if (arr19[i, 1].ToString().Contains("'")) arr19[i, 1] = arr19[i, 1].ToString().Replace('\'', ' ');

                        sql_comm_str = @"SELECT * FROM  [dbo].[uPhrenix] WHERE  [poz]='" + arr14[i, 1].ToString() + "' AND [date1]='" + arr15[i, 1].ToString() + "'";

                        comm2.CommandText = sql_comm_str;
                        SqlDataReader test_r1 = comm2.ExecuteReader();

                        if (test_r1.Read() == false)
                        {
                            test_r1.Close();
                            sql_comm_str = "INSERT [dbo].[uPhrenix] ([poz] ,[date1] ,[obj_name] ,[obj_addr] ,[temp1] ,[temp2]) VALUES ('" + arr14[i, 1].ToString() + "', '" + arr15[i, 1].ToString() + "', '" + arr16[i, 1].ToString() + "', '" + arr17[i, 1].ToString() + "', '" +"нет связи"+/* arr18[i, 1].ToString() */ "', '" + arr19[i, 1].ToString() + "')";

                            comm3.CommandText = sql_comm_str;
                            comm3.ExecuteNonQuery();
                            c++; //oooohhhh yeeaaaaaaaa 
                        }
                        else test_r1.Close();
                    }
                }
                toolStripProgressBar1.Value = 100;
                ObjWorkBookOut2.Close(null, null, null);
                MessageBox.Show("Добавлено тревог: " + c.ToString(), "Френикс, Тревоги --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        private void френиксToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;
            string sql_comm_str;

            sql_comm_str = @"SELECT * FROM  [dbo].[uPhrenix] WHERE [id]>1 ORDER BY [id]";

            comm2.CommandText = sql_comm_str;
            SqlDataReader test_r1 = comm2.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;
            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];

                rowcount++;
            }
            test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();

            //Создаём приложение.

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut = ObjExcel.Workbooks.Add(xls_path + "PritokA1.xls");

            //progressBar1.Value = 25;
            //label1.Text = "Run Excel App...";
            Refresh();

            object[,] arr1 = new object[rowcount, 29];
            object[,] arr3 = new object[rowcount, 29];
            int rowcount1 = rowcount;


            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[1];

            if (ws == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {
                for (int i = 0; i < rowcount; i++)
                {
                    arr1[i, 0] = ds1[i, 3].ToString();
                    arr1[i, 1] = ds1[i, 4].ToString();
                    arr1[i, 2] = ds1[i, 2].ToString();
                    arr1[i, 3] = ds1[i, 1].ToString();
                    arr1[i, 4] = ds1[i, 0].ToString();
                    arr1[i, 5] = ds1[i, 5].ToString();
                    arr1[i, 6] = ds1[i, 6].ToString();

                }

                int k = rowcount + 2;
                ws.get_Range("A" + 3, "AA" + k).Value2 = arr1;

            }

            /////////////////////////////////////////******************************************************/////////////////////////////////////////////

            sql_comm_str = @"SELECT [poz],[temp1]  ,COUNT([poz]) as c_poz
                            FROM [dbo].[uPhrenix]
                            WHERE [id]>1
                             GROUP BY [poz],[temp1] 
                            HAVING COUNT([poz]) > 1 " +
                " ORDER BY c_poz";

            comm2.CommandText = sql_comm_str;
            test_r1 = comm2.ExecuteReader();


            Refresh();

            object[,] ds2 = new object[150001, 4];
            rowcount = 0;
            while (test_r1.Read())
            {
                ds2[rowcount, 0] = test_r1[0];
                ds2[rowcount, 1] = test_r1[1];
                ds2[rowcount, 2] = test_r1[2];

                rowcount++;
            }
            test_r1.Close();

            object[,] arr2 = new object[rowcount, 10];

            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[3];

            if (ws2 == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {

                for (int i = 0; i < rowcount; i++)
                {
                    for (int j = 0; j < rowcount1; j++)
                    {
                        if (ds2[i, 0].ToString() == arr1[j, 3].ToString())
                        {
                            arr2[i, 0] = arr1[j, 0];
                            arr2[i, 1] = arr1[j, 1];
                            //[serv_org], [dogovor], [tel_info], [deblo_otvetish]
                            sql_comm_str = @"SELECT [serv_org],[deblo_otvetish],[tel_info] 
                            FROM [dbo].[uDisl]
                            WHERE [poz] LIKE '" + "С " + arr1[j, 3].ToString() + "'";

                            comm2.CommandText = sql_comm_str;
                            test_r1 = comm2.ExecuteReader();
                            if (test_r1.Read())
                            {

                                arr2[i, 5] = test_r1[0].ToString();
                                arr2[i, 6] = test_r1[1].ToString();
                                arr2[i, 7] = test_r1[2].ToString();

                                test_r1.Close();
                            }
                            else test_r1.Close();

                            break;
                        }
                    }

                    arr2[i, 2] = ds2[i, 0].ToString();
                    arr2[i, 4] = ds2[i, 1].ToString();
                    arr2[i, 3] = ds2[i, 2].ToString();

                }

                int m = 0;
                for (int n = 0; n < rowcount; n++)
                {
                    if (arr2[n, 4] != null)
                    {
                        arr2[m, 0] = arr2[n, 0];
                        arr2[m, 1] = arr2[n, 1];
                        arr2[m, 2] = arr2[n, 2];
                        arr2[m, 3] = arr2[n, 3];
                        arr2[m, 4] = arr2[n, 4];
                        arr2[m, 5] = arr2[n, 5];
                        arr2[m, 6] = arr2[n, 6];
                        arr2[m, 7] = arr2[n, 7];
                        m++;
                    }

                }
                int k = m + 2;
                ws2.get_Range("A" + 3, "I" + k).Value2 = arr2;

            }


            //progressBar1.Value = 100;
            //label1.Text = "Open excel window. Bye bye... ";
            Refresh();


            ObjExcel.Visible = true;

        }

        private void притокТестыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object[,] arr14 = new object[10001, 1];
            object[,] arr15 = new object[10001, 1];
            object[,] arr16 = new object[10001, 1];
            object[,] arr17 = new object[10001, 1];
            object[,] arr18 = new object[10001, 1];
            object[,] arr19 = new object[10001, 1];

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.ShowDialog(this);

            //MessageBox.Show("Путь к файлу: " + fdlg.FileName);

            toolStripProgressBar1.Value = 5;

            if (File.Exists(fdlg.FileName) && fdlg.FileName.Contains("\\\\") == false)
            {
                int c = 0;
                SqlCommand comm2 = new SqlCommand();
                comm2.Connection = Program.conn;
                SqlCommand comm3 = new SqlCommand();
                comm3.Connection = Program.conn;

                string sql_comm_str = "";


                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut2 = ObjExcel.Workbooks.Open(fdlg.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOut2;
                ObjWorkSheetOut2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut2.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range3 = ObjWorkSheetOut2.get_Range("E" + "1", "E" + "10000");
                arr14 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("G" + "1", "G" + "10000");
                arr15 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("I" + "1", "I" + "10000");
                arr16 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("J" + "1", "J" + "10000");
                arr17 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("L" + "1", "L" + "10000");
                arr18 = (object[,])range3.Value2;
                range3 = ObjWorkSheetOut2.get_Range("M" + "1", "M" + "10000");
                arr19 = (object[,])range3.Value2;

                for (int i = 1; i < 5000; i++)
                {
                    //progressBar1.Value++;
                    double progress1;
                    if (arr18[i, 1] == null) arr18[i, 1] = (object)" ";
                    if (arr19[i, 1] == null) arr19[i, 1] = (object)" ";

                    if (arr14[i, 1] != null && arr15[i, 1] != null /*&& arr14[i, 1].ToString() != "" && arr14[i, 1].ToString() != " "*/)
                    {
                        if (arr16[i, 1] == null) arr16[i, 1] = "(пустое поле)";
                        if (arr17[i, 1] == null) arr17[i, 1] = "(пустое поле)";

                        if (arr14[i, 1].ToString().Contains("'")) arr14[i, 1] = arr14[i, 1].ToString().Replace('\'', ' ');
                        if (arr15[i, 1].ToString().Contains("'")) arr15[i, 1] = arr15[i, 1].ToString().Replace('\'', ' ');
                        if (arr16[i, 1].ToString().Contains("'")) arr16[i, 1] = arr16[i, 1].ToString().Replace('\'', ' ');
                        if (arr17[i, 1].ToString().Contains("'")) arr17[i, 1] = arr17[i, 1].ToString().Replace('\'', ' ');
                        if (arr18[i, 1].ToString().Contains("'")) arr18[i, 1] = arr18[i, 1].ToString().Replace('\'', ' ');
                        if (arr19[i, 1].ToString().Contains("'")) arr19[i, 1] = arr19[i, 1].ToString().Replace('\'', ' ');

                        sql_comm_str = @"SELECT * FROM  [dbo].[uPritokGOLD] WHERE  [poz]='" + arr14[i, 1].ToString() + "' AND [date1]='" + arr15[i, 1].ToString() + "'";

                        comm2.CommandText = sql_comm_str;
                        SqlDataReader test_r1 = comm2.ExecuteReader();

                        if (test_r1.Read() == false)
                        {
                            test_r1.Close();
                            sql_comm_str = "INSERT [dbo].[uPritokGOLD] ([poz] ,[date1] ,[obj_name] ,[obj_addr] ,[temp1] ,[temp2]) VALUES ('" + arr14[i, 1].ToString() + "', '" + arr15[i, 1].ToString() + "', '" + arr16[i, 1].ToString() + "', '" + arr17[i, 1].ToString() + "', '" + arr18[i, 1].ToString() + "', '" + arr19[i, 1].ToString() + "')";

                            comm3.CommandText = sql_comm_str;
                            comm3.ExecuteNonQuery();
                            c++; //oooohhhh yeeaaaaaaaa 
                        }
                        else test_r1.Close();
                    }
                }
                toolStripProgressBar1.Value = 100;
                ObjWorkBookOut2.Close(null, null, null);
                MessageBox.Show("Добавлено тестов: " + c.ToString(), "Приток-А, Тесты --- Project Bonus 2024", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

        }

        private void притокТестыToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            List<string> ini_dates_yupi8 = new List<string>();
            if (File.Exists("pritonA.dates"))
            {
                ini_dates_yupi8.AddRange((string[])File.ReadAllLines("pritonA.dates", Encoding.GetEncoding(1251)));
            }

            //
            List<string> ini_date20241030_yupi8 = new List<string>();
            if (File.Exists("pritonA.20241030"))
            {
                ini_date20241030_yupi8.AddRange((string[])File.ReadAllLines("pritonA.20241030", Encoding.GetEncoding(1251)));
            }

            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;
            string sql_comm_str;

            sql_comm_str = @"SELECT * FROM  [dbo].[uPritokGOLD] WHERE [id]>1 ORDER BY [id]";

            comm2.CommandText = sql_comm_str;
            SqlDataReader test_r1 = comm2.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;
            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];

                rowcount++;
            }
            test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();

            SqlCommand comm3 = new SqlCommand();
            comm3.Connection = Program.conn;

            sql_comm_str = @"SELECT * FROM  [dbo].[uPritokA] WHERE [id]>1 ORDER BY [id]";

            comm3.CommandText = sql_comm_str;
            SqlDataReader test_r11 = comm3.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds11 = new object[150001, 51];
            int rowcount11 = 0;
            while (test_r11.Read())
            {
                char[] ch = new char[2];
                // = {' ','.'};
                ch[0] = ' ';
                ch[1] = '.';
                string[] strm = new string[10];
                strm = test_r11[2].ToString().Split(ch);


                if (strm.Length > 2)
                    ds11[rowcount11, 2] = strm[2] + "." + strm[1] + "." + strm[0];
                else
                    ds11[rowcount11, 2] = "неверные данные";


                ds11[rowcount11, 0] = test_r11[0];
                ds11[rowcount11, 1] = test_r11[1];
                
                ds11[rowcount11, 3] = test_r11[3];
                ds11[rowcount11, 4] = test_r11[4];
                ds11[rowcount11, 5] = test_r11[5];
                ds11[rowcount11, 6] = test_r11[6];

                rowcount11++;
            }
            test_r11.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();



            //Создаём приложение.

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut = ObjExcel.Workbooks.Add(xls_path + "PritokA1.xls");

            //progressBar1.Value = 25;
            //label1.Text = "Run Excel App...";
            Refresh();

            object[,] arr1 = new object[rowcount, 29];
            object[,] arr3 = new object[rowcount, 29];
            int rowcount1 = rowcount;


            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[1];

            if (ws == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {
                for (int i = 0; i < rowcount; i++)
                {
                    char[] ch = new char[2];
                    // = {' ','.'};
                    ch[0] = ' ';
                    ch[1] = '.';
                    string[] strm = new string[10];
                    strm = ds1[i, 2].ToString().Split(ch);


                    arr1[i, 0] = ds1[i, 3].ToString();
                    arr1[i, 1] = ds1[i, 4].ToString();
                    if (strm.Length > 2)
                        arr1[i, 2] = strm[2] + "." + strm[1] + "." + strm[0];//ds1[i, 2].ToString();
                    else
                        arr1[i, 2] = "неверные данные";
                    arr1[i, 3] = ds1[i, 1].ToString();
                    arr1[i, 4] = ds1[i, 0].ToString();

                    arr1[i, 5] = ds1[i, 5].ToString();
                    arr1[i, 6] = ds1[i, 6].ToString();

                }

                int k = rowcount + 2;
                ws.get_Range("A" + 3, "AA" + k).Value2 = arr1;

            }

            /////////////////////////////////////////******************************************************/////////////////////////////////////////////

            sql_comm_str = @"SELECT [poz],[temp1]  ,COUNT([poz]) as c_poz
                            FROM [dbo].[uPritokGOLD]
                            WHERE [id]>1
                             GROUP BY [poz],[temp1] 
                            HAVING COUNT([poz]) > 1 " +
                " ORDER BY c_poz";

            comm2.CommandText = sql_comm_str;
            test_r1 = comm2.ExecuteReader();


            Refresh();

            object[,] ds2 = new object[150001, 4];
            rowcount = 0;
            while (test_r1.Read())
            {
                ds2[rowcount, 0] = test_r1[0];
                ds2[rowcount, 1] = test_r1[1];
                ds2[rowcount, 2] = test_r1[2];

                rowcount++;
            }
            test_r1.Close();

            object[,] arr2 = new object[rowcount, 71];

            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[3];

            if (ws2 == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {

                for (int i = 0; i < rowcount; i++)
                {
                    /*if (!checkBox1.Checked && ds2[i, 1].ToString() == "Заявка")
                    {

                        continue;
                    }*/
                    for (int j = 0; j < rowcount1; j++)
                    {
                        if (ds2[i, 0].ToString() == arr1[j, 3].ToString())
                        {
                            arr2[i, 0] = arr1[j, 0];
                            arr2[i, 1] = arr1[j, 1];
                            //[serv_org], [dogovor], [tel_info], [deblo_otvetish]
                            sql_comm_str = @"SELECT [serv_org],[deblo_otvetish],[tel_info] 
                            FROM [dbo].[uDisl]
                            WHERE [poz] LIKE '" + "П " + arr1[j, 3].ToString() + "'";

                            comm2.CommandText = sql_comm_str;
                            test_r1 = comm2.ExecuteReader();
                            if (test_r1.Read())
                            {

                                arr2[i, 5] = test_r1[0].ToString();
                                arr2[i, 6] = test_r1[1].ToString();
                                arr2[i, 7] = test_r1[2].ToString();

                                test_r1.Close();
                            }
                            else test_r1.Close();
                            
                            int l = 0;
                            int p = 0;
                            for (l = 0; l < ini_dates_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount11; p++)
                                {
                                    if (/*arr1[j, 5].ToString() == "Тревога АН - авария направления" &&*/ arr1[j, 3].ToString() == ds11[p, 1].ToString() && ds11[p, 2].ToString() == ini_dates_yupi8[l])
                                    {
                                        arr2[i, 8 + l] = "н";
                                        break;
                                    }

                                }

                            }
                            

                            /*for (l = 0; l < ini_date20241030_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString() == ini_date20241030_yupi8[l])
                                    {
                                        arr2[i, 50] = "нет связи 30.10.2024!";
                                        break;
                                    }

                                }

                            }*/

                            break;
                        }
                    }

                    arr2[i, 2] = ds2[i, 0].ToString();
                    arr2[i, 4] = ds2[i, 1].ToString();
                    arr2[i, 3] = ds2[i, 2].ToString();

                }

                int m = 0;
                for (int n = 0; n < rowcount; n++)
                {
                    if (arr2[n, 4] != null)
                    {
                        arr2[m, 0] = arr2[n, 0];
                        arr2[m, 1] = arr2[n, 1];
                        arr2[m, 2] = arr2[n, 2];
                        arr2[m, 3] = arr2[n, 3];
                        arr2[m, 4] = arr2[n, 4];
                        arr2[m, 5] = arr2[n, 5];
                        arr2[m, 6] = arr2[n, 6];
                        arr2[m, 7] = arr2[n, 7];
                        m++;
                    }

                }
                int k = m + 2;
                ws2.get_Range("A" + 3, "BS" + k).Value2 = arr2;

                object[,] arr4 = new object[1, 71];
                for (int o = 0; o < ini_dates_yupi8.Count(); o++)
                {
                    arr4[0, o] = (object)ini_dates_yupi8[o];
                }
                ws2.get_Range("I" + 1, "BS" + 1).Value2 = arr4;
                

            }


            //progressBar1.Value = 100;
            //label1.Text = "Open excel window. Bye bye... ";
            Refresh();


            ObjExcel.Visible = true;

        }

        private void юпитерТестыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<string> ini_dates_yupi8 = new List<string>();
            if (File.Exists("yupi8.dates"))
            {
                ini_dates_yupi8.AddRange((string[])File.ReadAllLines("yupi8.dates", Encoding.GetEncoding(1251)));
            }

            //
            List<string> ini_date20241031_yupi8 = new List<string>();
            if (File.Exists("yupi8.20241031"))
            {
                ini_date20241031_yupi8.AddRange((string[])File.ReadAllLines("yupi8.20241031", Encoding.GetEncoding(1251)));
            }

            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = Program.conn;
            string sql_comm_str;

            sql_comm_str = @"SELECT * FROM  [dbo].[uJupiterGOLD] WHERE [id]>1 ORDER BY [id]";

            comm2.CommandText = sql_comm_str;
            SqlDataReader test_r1 = comm2.ExecuteReader();


            progressBar1.Value = 10;
            label1.Text = "SQL query...";
            Refresh();
            
            object[,] ds1 = new object[150001, 51];
            int rowcount = 0;
            while (test_r1.Read())
            {
                ds1[rowcount, 0] = test_r1[0];
                ds1[rowcount, 1] = test_r1[1];
                ds1[rowcount, 2] = test_r1[2];
                ds1[rowcount, 3] = test_r1[3];
                ds1[rowcount, 4] = test_r1[4];
                ds1[rowcount, 5] = test_r1[5];
                ds1[rowcount, 6] = test_r1[6];

                rowcount++;
            }
            test_r1.Close();///////////////**********************************************/////////////////////////////////////////////////////

            //progressBar1.Value = 15;
            //label1.Text = "Copy data...";
            Refresh();

            SqlCommand comm3 = new SqlCommand();
            comm3.Connection = Program.conn;

            sql_comm_str = @"SELECT * FROM  [dbo].[uJupiter8] WHERE [id]>1 ORDER BY [id]";

            comm3.CommandText = sql_comm_str;
            SqlDataReader test_r11 = comm3.ExecuteReader();


            //progressBar1.Value = 10;
            //label1.Text = "SQL query...";
            Refresh();


            object[,] ds11 = new object[150001, 51];
            int rowcount11 = 0;
            while (test_r11.Read())
            {
                ds11[rowcount11, 0] = test_r11[0];
                ds11[rowcount11, 1] = test_r11[1];
                ds11[rowcount11, 2] = test_r11[2];
                ds11[rowcount11, 3] = test_r11[3];
                ds11[rowcount11, 4] = test_r11[4];
                ds11[rowcount11, 5] = test_r11[5];
                ds11[rowcount11, 6] = test_r11[6];

                rowcount11++;
            }
            test_r11.Close();///////////////**********************************************/////////////////////////////////////////////////////

            progressBar1.Value = 15;
            label1.Text = "Copy data...";
            Refresh();

            //Создаём приложение.

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOut = ObjExcel.Workbooks.Add(xls_path + "PritokA1.xls");

            progressBar1.Value = 25;
            label1.Text = "Run Excel App...";
            Refresh();

            object[,] arr1 = new object[rowcount, 29];
            object[,] arr3 = new object[rowcount, 29];
            int rowcount1 = rowcount;


            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[1];

            if (ws == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {
                for (int i = 0; i < rowcount; i++)
                {
                    char[] ch = new char[2];
                    // = {' ','.'};
                    ch[0] = ' ';
                    ch[1] = '.';
                    string[] strm = new string[10];
                    strm = ds1[i, 2].ToString().Split(ch);


                    arr1[i, 0] = ds1[i, 3].ToString();
                    arr1[i, 1] = ds1[i, 4].ToString();
                    if (strm.Length > 2)
                        arr1[i, 2] = strm[2] + "." + strm[1] + "." + strm[0];//ds1[i, 2].ToString();
                    else
                        arr1[i, 2] = "неверные данные";
                    arr1[i, 3] = ds1[i, 1].ToString();
                    arr1[i, 4] = ds1[i, 0].ToString();

                    arr1[i, 5] = ds1[i, 5].ToString();
                    arr1[i, 6] = ds1[i, 6].ToString();

                }

                int k = rowcount + 2;
                ws.get_Range("A" + 3, "AA" + k).Value2 = arr1;

            }

            /////////////////////////////////////////******************************************************/////////////////////////////////////////////

            sql_comm_str = @"SELECT [poz],[temp1]  ,COUNT([poz]) as c_poz
                            FROM [dbo].[uJupiterGOLD]
                            WHERE [id]>1
                             GROUP BY [poz],[temp1] 
                            HAVING COUNT([poz]) > 1 " +
                " ORDER BY c_poz";

            comm2.CommandText = sql_comm_str;
            test_r1 = comm2.ExecuteReader();


            Refresh();

            object[,] ds2 = new object[150001, 4];
            rowcount = 0;
            while (test_r1.Read())
            {
                ds2[rowcount, 0] = test_r1[0];
                ds2[rowcount, 1] = test_r1[1];
                ds2[rowcount, 2] = test_r1[2];

                rowcount++;
            }
            test_r1.Close();

            object[,] arr2 = new object[rowcount, 71];

            Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBookOut.Worksheets[3];

            if (ws2 == null)
            {
                //("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            else
            {

                for (int i = 0; i < rowcount; i++)
                {
                    /*if (!checkBox1.Checked && ds2[i, 1].ToString() == "Заявка")
                    {

                        continue;
                    }*/

                    for (int j = 0; j < rowcount1; j++)
                    {
                        if (ds2[i, 0].ToString() == arr1[j, 3].ToString())
                        {
                            arr2[i, 0] = arr1[j, 0];
                            arr2[i, 1] = arr1[j, 1];
                            //[serv_org], [dogovor], [tel_info], [deblo_otvetish]
                            sql_comm_str = @"SELECT [serv_org],[deblo_otvetish],[tel_info] 
                            FROM [dbo].[uDisl]
                            WHERE [poz] LIKE '" + "Ю " + arr1[j, 3].ToString() + "'";

                            comm2.CommandText = sql_comm_str;
                            test_r1 = comm2.ExecuteReader();
                            if (test_r1.Read())
                            {

                                arr2[i, 5] = test_r1[0].ToString();
                                arr2[i, 6] = test_r1[1].ToString();
                                arr2[i, 7] = test_r1[2].ToString();

                                test_r1.Close();
                            }
                            else test_r1.Close();

                            int l = 0;
                            int p = 0;
                            for (l = 0; l < ini_dates_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount11; p++)
                                {
                                    if (arr1[j, 3].ToString() == ds11[p, 1].ToString() && ds11[p, 2].ToString() == ini_dates_yupi8[l])
                                    {
                                        arr2[i, 8 + l] = "н";
                                        break;
                                    }

                                }

                            }
                            
                            char[] ch = new char[1];
                            ch[0] = '0';

                            /*for (l = 0; l < ini_date20241031_yupi8.Count(); l++)
                            {
                                for (p = 0; p < rowcount1; p++)
                                {
                                    if (arr1[j, 3].ToString() == arr1[p, 3].ToString() && arr1[p, 3].ToString().Trim(ch) == ini_date20241031_yupi8[l])
                                    {
                                        arr2[i, 50] = "нет связи 31.10.2024!";
                                        break;
                                    }

                                }

                            }*/

                            break;
                        }
                    }

                    arr2[i, 2] = ds2[i, 0].ToString();
                    arr2[i, 4] = ds2[i, 1].ToString();
                    arr2[i, 3] = ds2[i, 2].ToString();

                }

                int m = 0;
                for (int n = 0; n < rowcount; n++)
                {
                    if (arr2[n, 4] != null)
                    {
                        arr2[m, 0] = arr2[n, 0];
                        arr2[m, 1] = arr2[n, 1];
                        arr2[m, 2] = arr2[n, 2];
                        arr2[m, 3] = arr2[n, 3];
                        arr2[m, 4] = arr2[n, 4];
                        arr2[m, 5] = arr2[n, 5];
                        arr2[m, 6] = arr2[n, 6];
                        arr2[m, 7] = arr2[n, 7];
                        m++;
                    }

                }
                int k = m + 2;
                ws2.get_Range("A" + 3, "BS" + k).Value2 = arr2;

                object[,] arr4 = new object[1, 71];
                for (int o = 0; o < ini_dates_yupi8.Count(); o++)
                {
                    arr4[0, o] = (object)ini_dates_yupi8[o];
                }
                ws2.get_Range("I" + 1, "BS" + 1).Value2 = arr4;
                

            }


            //progressBar1.Value = 100;
            //label1.Text = "Open excel window. Bye bye... ";
            Refresh();


            ObjExcel.Visible = true;


        }

        private void отчетыToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
