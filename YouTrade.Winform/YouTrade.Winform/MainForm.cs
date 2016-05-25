﻿using Excel;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YouTrade.Winform
{
    public partial class MainForm : Form
    {
        string sqlConnectionString = ConfigurationManager.AppSettings["connectionString"];

        string pathIn = "", pathOut = "", pathTempIncome = "", pathTempBasicInfo = "";
        DataSet dsSource = null;
        DataSet dsSource1 = null;
        int demIncome = 0, demBasicInfo = 0;

        private void btnInput_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            DialogResult result = fbd.ShowDialog();
            if(result==DialogResult.OK)
                tbInput.Text = fbd.SelectedPath + "\\";
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            DialogResult result = fbd.ShowDialog();
            if(result==DialogResult.OK)
                tbOutput.Text = fbd.SelectedPath + "\\";
        }

        public MainForm()
        {
            InitializeComponent();

            // Path Input
            pathIn = System.Windows.Forms.Application.StartupPath + "\\Input\\";
            if (!Directory.Exists(pathIn))
            {
                Directory.CreateDirectory(pathIn);
            }
            tbInput.Text = pathIn;

            // Path Output
            string pathOut1 = System.Windows.Forms.Application.StartupPath + "\\Output";
            pathOut = pathOut1 + "\\";
            if (!Directory.Exists(pathOut))
            {
                Directory.CreateDirectory(pathOut);
            }
            tbOutput.Text = pathOut1;

            // Path TempIncome
            pathTempIncome = System.Windows.Forms.Application.StartupPath + "\\Output\\Income\\Temp\\";
            if (!Directory.Exists(pathTempIncome))
            {
                Directory.CreateDirectory(pathTempIncome);
            }

        }

        private void Click_Ratios(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            btnIncome.Text = "Income Running...";


            MoveToTempIncome();
            ReadExcelAndSaveIncome();



            btnIncome.Text = "Income";
        }
        void ReadExcelAndSaveIncome()
        {
            try
            {
                var files1 = Directory.GetFiles(pathTempIncome, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls"));
                foreach (string file in files1)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);

                        string fullNameIn_In_Out = tbOutput.Text + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {

                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBIncome(tbl);
                                break;
                            }

                        }
                        StoreFileIncome(file);
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
               

            }
        }
        // Move to temp and chang xls Income
        void MoveToTempIncome()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("income") && !f.Contains("~$"));

            Microsoft.Office.Interop.Excel.Application excelApp = null;// = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;

            foreach (string file in files)
            {

                try
                {
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.FileValidation = MsoFileValidationMode.msoFileValidationSkip;

                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string fileEx = Path.GetExtension(file);

                    string FullNameIn = tbInput.Text + fileName + fileEx;
                    string fullNameIn_In_Temp = pathTempIncome + fileName.Replace(".", string.Empty) + ".xls";
                    if (!File.Exists(fullNameIn_In_Temp))
                    {
                        excelWorkbook = excelApp.Workbooks.Open(FullNameIn, 1, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, null, false);

                        excelApp.DisplayAlerts = false;
                        string fileNameOut = pathTempIncome + fileName.Replace(".", string.Empty);//+ "-" + DateTime.Now.Hour + "-" + DateTime.Now.Minute + "";

                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    //if (File.Exists(FullNameIn))
                    //{
                    //    File.Delete(FullNameIn);
                    //}
                }
                catch (Exception ex)
                {
                    /////   MessageBox.Show(ex.ToString());
                 //   listBox2.Items.Add("MoveToTemp: " + file);
                }
                finally
                {
                    if (excelWorkbook != null)
                    {
                        Marshal.FinalReleaseComObject(excelWorkbook);
                        excelWorkbook = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.FinalReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
            }
        }
        private void StoreFileIncome(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);

                if (!Directory.Exists(tbOutput.Text + "\\Income\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\Income\\");
                }

                File.Copy(fileName, tbOutput.Text+"\\Income\\" + filenameOnly, true);
                File.Delete(fileName);
            }

        }
        private void SaveToDBIncome(System.Data.DataTable dt)
        {
            //string[] arr = new string[3];
            //arr[0] = "Year:";
            //arr[1] = "Quarter:";
            //arr[2] = "Unit:";
            //string[] arrOut;
            List<string> listIDFeild = new List<string>();
            int dong = 0, cot = 0;
            //demIncome = (dt.Columns.Count - 5) * (dt.Rows.Count - 9);
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();
                // Save Feild
                for (int i = 4; i <= dt.Columns.Count - 1; i++)
                {
                    string strQueryFeild = "IF NOT EXISTS (SELECT 1 FROM DBO.Income_Feild WHERE IDfeild=@idfeild ) BEGIN INSERT INTO [dbo].[Income_Feild] ([IDFeild],[unit]) VALUES (@idfeild,@unit) END";
                    SqlCommand sqlcmd = new SqlCommand(strQueryFeild, dbcon);
                    string pattern = dt.Rows[6][i].ToString();
                    if (pattern.Trim() == "")
                        break;
                    //  string[] st = pattern.Split(new string[] { "Consolidated Year:", "Quarter:", "Unit:" }, StringSplitOptions.RemoveEmptyEntries);
                    // print array st

                    // Get Name
                    int startPositionName = pattern.IndexOf(".") + ".".Length;
                    string patternT = pattern.Substring(startPositionName, pattern.Length - startPositionName);

                    int startPositionNameDot = patternT.IndexOf(".") + ".".Length;
                    string patternTDot = patternT.Substring(startPositionNameDot, patternT.Length - startPositionNameDot);

                    string name = patternTDot.Substring(0, patternTDot.IndexOf("Consolidated\nYear:"));
                    // Get Unit
                    int startUnitPosition = patternTDot.IndexOf("Unit:") + "Unit:".Length;
                    string unit = patternTDot.Substring(startUnitPosition, patternTDot.Length - startUnitPosition);

                    sqlcmd.Parameters.AddWithValue("@idfeild", name.Trim());
                    sqlcmd.Parameters.AddWithValue("@unit", unit.Trim());

                    sqlcmd.ExecuteNonQuery();
                    cot++;
                }



                // Save Income_Financial
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    try
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        // Save to Finance
                        string strQueryFinance = "IF NOT EXISTS (SELECT 1 FROM DBO.Income_Financial WHERE Ticker=@Ticker ) BEGIN INSERT INTO [dbo].[Income_Financial] ([Ticker],[Name],[Exchange]) VALUES (@ticker,@name,@exchange) END";
                        SqlCommand sqlcmd = new SqlCommand(strQueryFinance, dbcon);
                        sqlcmd.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                        sqlcmd.Parameters.AddWithValue("@name", dt.Rows[i][2].ToString().Trim());
                        sqlcmd.Parameters.AddWithValue("@exchange", dt.Rows[i][3].ToString().Trim());
                        sqlcmd.ExecuteNonQuery();
                        dong++;
                    }
                    catch
                    {
                        //listBox2.Items.Add("SaveIncomeFinancial: " + i.ToString().Trim());
                    }
                }

                demIncome += dong * cot;
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    try
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        for (int j = 4; j <= dt.Columns.Count - 1; j++)
                        {

                            string strQueryDetails = "INSERT INTO [dbo].[Income_Details_Feild]([Ticker] ,[IDFeild] ,[Year],[IDQuarter],[Value]) VALUES(@ticker,@explore,@year,@quarter,@value)";
                            SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);
                            sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                            string pattern = dt.Rows[6][j].ToString();
                            // string[] st = pattern.Split(new string[] { "Year:", "Quarter:", "Unit:" }, StringSplitOptions.RemoveEmptyEntries);
                            // print array st
                            // Get Name
                            int startPositionName = pattern.IndexOf(".") + ".".Length;
                            string patternT = pattern.Substring(startPositionName, pattern.Length - startPositionName);

                            int startPositionNameDot = patternT.IndexOf(".") + ".".Length;
                            string patternTDot = patternT.Substring(startPositionNameDot, patternT.Length - startPositionNameDot);

                            string name = patternTDot.Substring(0, patternTDot.IndexOf("Consolidated\nYear:"));

                            // Get Year
                            int startPositionYear = patternTDot.IndexOf("Consolidated\nYear:") + "Consolidated\nYear:".Length;
                            string year = patternTDot.Substring(startPositionYear, patternTDot.IndexOf("Quarter") - startPositionYear);

                            // Get Quarter
                            int startPositionQuarter = patternTDot.IndexOf("Quarter") + "Quarter:".Length;
                            string quarter = patternTDot.Substring(startPositionQuarter, patternTDot.IndexOf("Unit") - startPositionQuarter);

                            sqlcmdD.Parameters.AddWithValue("@explore", name.ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(year));
                            sqlcmdD.Parameters.AddWithValue("@quarter", quarter.ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@value", dt.Rows[i][j].ToString().Trim());
                            sqlcmdD.ExecuteNonQuery();
                            if (dt.Rows[i][1].ToString().Trim() == "")
                                break;
                        }
                    }
                    catch
                    {
                      //  listBox2.Items.Add("SaveValueIncome: " + i.ToString() + "-");
                    }
                }
                dbcon.Close();
            }
         
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }
        DataSet GetDatasetFromExcel(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            if (excelReader == null)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);


            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            return result;

        }
    }
}
