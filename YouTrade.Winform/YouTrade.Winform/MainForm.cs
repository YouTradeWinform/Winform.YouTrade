using Excel;
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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YouTrade.Winform
{
    public partial class MainForm : Form
    {
        string sqlConnectionString = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;

        string pathIn = "", pathOut = "";
        string pathTempIncome = "", pathTempBasicInfo = "";
        string pathTempRatios = "", pathTempBalance = "", pathTempStock = "";
        string pathTempCashFlow = "", pathTempNote = "";
        DataSet dsSource = null;
        DataSet dsSource1 = null;
        public int demFileRatios = 0, demFileBalance = 0, demFileStock = 0;
        int demIncome = 0, demBasicInfo = 0;
        bool KTIncome = false;

        #region Click_Input_Output
        private void Click_Input(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
                tbInput.Text = fbd.SelectedPath + "\\";
        }
        private void Click_Output(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
                tbOutput.Text = fbd.SelectedPath + "\\";
        }
        #endregion

        #region Start Main, create foler for each type
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
            //Path Temp Ratios
            pathTempRatios = System.Windows.Forms.Application.StartupPath + "\\Output\\Ratios\\Temp\\";
            if (!Directory.Exists(pathTempRatios))
            {
                Directory.CreateDirectory(pathTempRatios);
            }
            //Path Temp Balance
            pathTempBalance = System.Windows.Forms.Application.StartupPath + "\\Output\\Balance\\Temp\\";
            if (!Directory.Exists(pathTempBalance))
            {
                Directory.CreateDirectory(pathTempBalance);
            }
            //Path Temp Stock Market Data
            pathTempStock = System.Windows.Forms.Application.StartupPath + "\\Output\\Stock\\Temp\\";
            if (!Directory.Exists(pathTempStock))
            {
                Directory.CreateDirectory(pathTempStock);
            }
            //Path Temp CashFlow
            pathTempCashFlow = System.Windows.Forms.Application.StartupPath + "\\Output\\CashFlow\\Temp\\";
            if (!Directory.Exists(pathTempCashFlow))
            {
                Directory.CreateDirectory(pathTempCashFlow);
            }
            //Path Temp Note
            pathTempNote = System.Windows.Forms.Application.StartupPath + "\\Output\\Note\\Temp\\";
            if (!Directory.Exists(pathTempNote))
            {
                Directory.CreateDirectory(pathTempNote);
            }
        }
        #endregion

        #region 7 click
        private void Click_Ratios(object sender, EventArgs e)
        {
            btnRatios.Text = "Running...";
            btnRatios.BackColor = Color.Gray;
            txtStatus.Text = "Loading...";
            MoveToTempRatios();
            //progressBar1.Minimum = 0;
            //progressBar1.Maximum = demFileRatios;
            //progressBar1.Value = 1;
            //progressBar1.Step = 10;
            ReadExcelAndSaveRatios();
            txtStatus.Text = "Done!";
            btnRatios.Text = "Ratios";
            btnRatios.BackColor = Color.White;
        }
        private void Click_Balance(object sender, EventArgs e)
        {
            btnBalance.Text = "Running...";
            btnBalance.BackColor = Color.Gray;
            txtStatus.Text = "Loading...";
            MoveToTempBalance();
            //progressBar1.Minimum = 0;
            //progressBar1.Maximum = demFileBalance;
            //progressBar1.Value = 1;
            //progressBar1.Step = 10;
            ReadExcelAndSaveBalance();
            txtStatus.Text = "Done!";
            btnBalance.Text = "Balance";
            btnRatios.BackColor = Color.White;
        }
        private void Click_Stock(object sender, EventArgs e)
        {
            btnStock.Text = "Running...";
            btnStock.BackColor = Color.Gray;
            txtStatus.Text = "Loading...";
            MoveToTempStock();
            //progressBar1.Minimum = 0;
            //progressBar1.Maximum = demFileStock;
            //progressBar1.Value = 1;
            //progressBar1.Step = 10;
            ReadExcelAndSaveStock();
            txtStatus.Text = "Done!";
            btnStock.Text = "Stock";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            btnIncome.Text = "Income Running...";



            MoveToTempIncome();
            txtStatus.Text = "Done!";
            progressBar1.Value = 0;

            ReadExcelAndSaveIncome();

            CheckIfFileInTempIncome();

            btnIncome.Text = "Income";
        }
        void CheckIfFileInTempIncome()
        {


        }
        #endregion

        #region GetDatasetFromExcel
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
        #endregion

        #region Move file to temp
        //Income
        void MoveToTempIncome()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("income") && !f.Contains("~$"));
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            // progressBar1 = new ProgressBar();
            //progressBar1.Value = 0; // progressbar
            // progressBar1.Maximum = files.Count(); // progressbar
            //   progressBar1.Step = 1; // progressbar.





            foreach (string file in files)
            {
                // progressBar1.Value++; // progressbar
                txtStatus.Text = "Moving... " + Path.GetFileName(file); // progressbar
                progressBar1.Maximum = files.Count();
                progressBar1.Step = 1;

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
                        string fileNameOut = pathTempIncome + fileName.Replace(".", string.Empty);
                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    //if (File.Exists(FullNameIn))
                    //{
                    //    File.Delete(FullNameIn);
                    //}
                }
                catch (Exception ex)
                {
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
                progressBar1.Value++; // progressbar


            }
            //  progressBar1.Value= progressBar1.Maximum;
            //Thread.Sleep(1000);
        }

        //Ratios
        public void MoveToTempRatios()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("ratio") && !f.Contains("~$"));
            Microsoft.Office.Interop.Excel.Application excelApp = null;
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
                    string fullNameIn_In_Temp = pathTempRatios + fileName.Replace(".", string.Empty) + ".xls";
                    if (!File.Exists(fullNameIn_In_Temp))
                    {
                        excelWorkbook = excelApp.Workbooks.Open(FullNameIn, 1, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, null, false);
                        excelApp.DisplayAlerts = false;
                        string fileNameOut = pathTempRatios + fileName.Replace(".", string.Empty);
                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    demFileRatios = demFileRatios + 1;
                }
                catch (Exception ex)
                {
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
        //Balance
        public void MoveToTempBalance()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("balance") && !f.Contains("~$"));
            Microsoft.Office.Interop.Excel.Application excelApp = null;
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
                    string fullNameIn_In_Temp = pathTempBalance + fileName.Replace(".", string.Empty) + ".xls";
                    if (!File.Exists(fullNameIn_In_Temp))
                    {
                        excelWorkbook = excelApp.Workbooks.Open(FullNameIn, 1, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, null, false);
                        excelApp.DisplayAlerts = false;
                        string fileNameOut = pathTempBalance + fileName.Replace(".", string.Empty);
                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    demFileBalance = demFileBalance + 1;
                }
                catch (Exception ex)
                {
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
        //Stock
        public void MoveToTempStock()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("StockMarketData") && !f.Contains("~$"));
            Microsoft.Office.Interop.Excel.Application excelApp = null;
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
                    string fullNameIn_In_Temp = pathTempStock + fileName.Replace(".", string.Empty) + ".xls";
                    if (!File.Exists(fullNameIn_In_Temp))
                    {
                        excelWorkbook = excelApp.Workbooks.Open(FullNameIn, 1, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, null, false);
                        excelApp.DisplayAlerts = false;
                        string fileNameOut = pathTempStock + fileName.Replace(".", string.Empty);
                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    demFileStock = demFileStock + 1;
                }
                catch (Exception ex)
                {
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
        #endregion
        #region Move file to temp
        //CashFlow
        void MoveToTempCashFlow()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("cashflow") && !f.Contains("~$"));
            Microsoft.Office.Interop.Excel.Application excelApp = null;
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
                    string fullNameIn_In_Temp = pathTempCashFlow + fileName.Replace(".", string.Empty) + ".xls";
                    if (!File.Exists(fullNameIn_In_Temp))
                    {
                        excelWorkbook = excelApp.Workbooks.Open(FullNameIn, 1, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, null, false);
                        excelApp.DisplayAlerts = false;
                        string fileNameOut = pathTempCashFlow + fileName.Replace(".", string.Empty);
                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                }
                catch (Exception ex)
                {
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
        //Note
        void MoveToTempNote()
        {
            // Chuyển file sang folder Temp
            var files = Directory.GetFiles(tbInput.Text, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsm") || s.EndsWith(".xlsx")).Where(f => f.Contains("note") && !f.Contains("~$"));
            Microsoft.Office.Interop.Excel.Application excelApp = null;
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
                    string fullNameIn_In_Temp = pathTempNote + fileName.Replace(".", string.Empty) + ".xls";
                    if (!File.Exists(fullNameIn_In_Temp))
                    {
                        excelWorkbook = excelApp.Workbooks.Open(FullNameIn, 1, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, null, false);
                        excelApp.DisplayAlerts = false;
                        string fileNameOut = pathTempNote + fileName.Replace(".", string.Empty);
                        excelWorkbook.SaveAs(fileNameOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                }
                catch (Exception ex)
                {
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

        #endregion
        #region Save to DB
        //Ratios
        void SaveToDBRatio(System.Data.DataTable dt)
        {
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();
                //Ratios
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    for (int j = 4; j <= dt.Columns.Count - 1; j++)
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        string strQueryDetails = "INSERT INTO [dbo].[Ratio]([Ticker],[Year],[Quater],[Name],[Value],[Unit]) VALUES(@ticker,@year,@quater,@explorename,@value,@unit)";
                        SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);

                        string pattern = dt.Rows[6][j].ToString();
                        string[] st = pattern.Split(new string[] { "Year:", "Quarter:", "Unit:" }, StringSplitOptions.RemoveEmptyEntries);
                        //Name
                        string name = pattern.Substring(0, pattern.IndexOf("\nTrailing"));
                        //Year
                        int startPositionYear = pattern.IndexOf("Year:") + "Year:".Length;
                        string year = pattern.Substring(startPositionYear, pattern.IndexOf("\nQuarter") - startPositionYear);
                        int years = int.Parse(year);
                        //Quater
                        int startPositionQuarter = pattern.IndexOf("Quarter") + "Quarter:".Length;
                        string quarter = pattern.Substring(startPositionQuarter, pattern.IndexOf("\nUnit") - startPositionQuarter);
                        if (quarter == " Annual")
                            quarter = "5";
                        //Unit
                        int startUnitPosition = pattern.IndexOf("Unit:") + "Unit:".Length;
                        string unit = pattern.Substring(startUnitPosition, pattern.Length - startUnitPosition);
                        if (true == (unit.Contains("\n")))
                            unit = unit.Substring(0, unit.Length);
                        string value = dt.Rows[i][j].ToString().Trim();
                        if(value=="")
                        {
                            string strQueryDetails1 = "INSERT INTO [dbo].[Ratio]([Ticker],[Year],[Quater],[Name],[Unit]) VALUES(@ticker,@year,@quater,@explorename,@unit)";
                            SqlCommand sqlcmdD1 = new SqlCommand(strQueryDetails1, dbcon);
                            sqlcmdD1.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                            sqlcmdD1.Parameters.AddWithValue("@explorename", name.ToString().Trim());
                            sqlcmdD1.Parameters.AddWithValue("@year", Convert.ToInt16(years));
                            sqlcmdD1.Parameters.AddWithValue("@quater", Convert.ToInt16(quarter));
                            sqlcmdD1.Parameters.AddWithValue("@unit", unit.ToString().Trim());
                            sqlcmdD1.ExecuteNonQuery();
                        }
                        else
                        { 
                            sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@explorename", name.ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(years));
                            sqlcmdD.Parameters.AddWithValue("@quater", Convert.ToInt16(quarter));
                            sqlcmdD.Parameters.AddWithValue("@value", double.Parse(value));
                            sqlcmdD.Parameters.AddWithValue("@unit", unit.ToString().Trim());
                            sqlcmdD.ExecuteNonQuery();
                        }
                    }
                }
                dbcon.Close();
            }
        }
        // Save to balance
        void SaveToDBBalance(System.Data.DataTable dt)
        {
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    for (int j = 4; j <= dt.Columns.Count - 1; j++)
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        string strQueryDetails = "INSERT INTO [dbo].[BalanceSheet]([Ticker],[Year],[Quater],[Name],[Value],[Unit]) VALUES(@ticker,@year,@quater,@explorename,@value,@unit)";
                        SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);

                        string pattern = dt.Rows[6][j].ToString();
                        string[] st = pattern.Split(new string[] { "Year:", "Quarter:", "Unit:" }, StringSplitOptions.RemoveEmptyEntries);
                        //Name
                        string name = pattern.Substring(0, pattern.IndexOf("\nConsolidated"));
                        //Year
                        int startPositionYear = pattern.IndexOf("Year:") + "Year:".Length;
                        string year = pattern.Substring(startPositionYear, pattern.IndexOf("\nQuarter") - startPositionYear);
                        int years = int.Parse(year);
                        //Quater
                        int startPositionQuarter = pattern.IndexOf("Quarter") + "Quarter:".Length;
                        string quarter = pattern.Substring(startPositionQuarter, pattern.IndexOf("\nUnit") - startPositionQuarter);
                        if (quarter == " Annual")
                            quarter = "5";
                        //Unit
                        int startUnitPosition = pattern.IndexOf("Unit:") + "Unit:".Length;
                        string unit = pattern.Substring(startUnitPosition, pattern.Length - startUnitPosition);
                        if (true == (unit.Contains("\n")))
                            unit = unit.Substring(0, unit.Length);

                        sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@explorename", name.ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(years));
                        sqlcmdD.Parameters.AddWithValue("@quater", Convert.ToInt16(quarter));
                        sqlcmdD.Parameters.AddWithValue("@value", dt.Rows[i][j].ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@unit", unit.ToString().Trim());
                        sqlcmdD.ExecuteNonQuery();
                    }
                }
                dbcon.Close();
            }
        }
        //Stock
        void SaveToDBStock(System.Data.DataTable dt)
        {
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();
                if (dt.Columns.Count > 12)
                {
                    for (int i = 8; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][0].ToString().Trim() == "")
                            break;
                        string close = dt.Rows[i][5].ToString().Trim();
                        string closeadj = dt.Rows[i][6].ToString().Trim();
                        string highet = dt.Rows[i][22].ToString().Trim();
                        string highetsadj = dt.Rows[i][23].ToString().Trim();
                        string lowest = dt.Rows[i][24].ToString().Trim();
                        string lowestadj = dt.Rows[i][25].ToString().Trim();
                        string open = dt.Rows[i][26].ToString().Trim();
                        string openadj = dt.Rows[i][27].ToString().Trim();
                        string volume = dt.Rows[i][40].ToString().Trim();
                        if (close == "")
                        {
                            if (closeadj == "")
                            {
                                if (highet == "")
                                {
                                    if (highetsadj == "")
                                    {
                                        if (lowest == "")
                                        {
                                            if (lowestadj == "")
                                            {
                                                if (open == "")
                                                {
                                                    if (openadj == "")
                                                    {
                                                        if (volume == "")
                                                        {
                                                            // Ticker, Date
                                                            string strQueryDetails0 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ) VALUES(@ticker,@trading)";
                                                            SqlCommand sqlcmdD0 = new SqlCommand(strQueryDetails0, dbcon);
                                                            sqlcmdD0.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                                            sqlcmdD0.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                                            sqlcmdD0.ExecuteNonQuery();
                                                        }
                                                        else
                                                        {
                                                            //Tiker, Date, Volume
                                                            string strQueryDetails1 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Volume]) VALUES(@ticker,@trading, @TotaTradingVolumes)";
                                                            SqlCommand sqlcmdD1 = new SqlCommand(strQueryDetails1, dbcon);
                                                            sqlcmdD1.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                                            sqlcmdD1.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                                            sqlcmdD1.Parameters.AddWithValue("@TotaTradingVolumes",double.Parse(volume));
                                                            sqlcmdD1.ExecuteNonQuery();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //Tiker, Date, Volume, OpenAdj
                                                        string strQueryDetails2 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[OpenAdjusted],[Volume]) VALUES(@ticker,@trading, @OpenAdjusted, @TotaTradingVolumes)";
                                                        SqlCommand sqlcmdD2 = new SqlCommand(strQueryDetails2, dbcon);
                                                        sqlcmdD2.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                                        sqlcmdD2.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                                        sqlcmdD2.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                                        sqlcmdD2.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                                        sqlcmdD2.ExecuteNonQuery();
                                                    }
                                                }
                                                else
                                                {
                                                    //Tiker, Date, Volume, OpenAdj, Open
                                                    string strQueryDetails3 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date],[Open],[OpenAdjusted],[Volume]) VALUES(@ticker,@trading,@Opens, @OpenAdjusted, @TotaTradingVolumes)";
                                                    SqlCommand sqlcmdD3 = new SqlCommand(strQueryDetails3, dbcon);
                                                    sqlcmdD3.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                                    sqlcmdD3.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                                    sqlcmdD3.Parameters.AddWithValue("@Opens", double.Parse(open));
                                                    sqlcmdD3.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                                    sqlcmdD3.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                                    sqlcmdD3.ExecuteNonQuery();
                                                }
                                            }
                                            else
                                            {
                                                //Tiker, Date, Volume, OpenAdj, Open, LowestAdj
                                                string strQueryDetails4 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Open],[OpenAdjusted],[LowestAdjusted],[Volume]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted, @LowestAdjusted, @TotaTradingVolumes)";
                                                SqlCommand sqlcmdD4 = new SqlCommand(strQueryDetails4, dbcon);
                                                sqlcmdD4.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                                sqlcmdD4.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                                sqlcmdD4.Parameters.AddWithValue("@LowestAdjusted", double.Parse(lowestadj));
                                                sqlcmdD4.Parameters.AddWithValue("@Opens", double.Parse(open));
                                                sqlcmdD4.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                                sqlcmdD4.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                                sqlcmdD4.ExecuteNonQuery();
                                            }
                                        }
                                        else
                                        {
                                            //Tiker, Date, Volume, OpenAdj, Open, LowestAdj, Lowest
                                            string strQueryDetails5 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Open],[OpenAdjusted],[Lowest],[LowestAdjusted],[Volume]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted,  @Lowest, @LowestAdjusted, @TotaTradingVolumes)";
                                            SqlCommand sqlcmdD5 = new SqlCommand(strQueryDetails5, dbcon);
                                            sqlcmdD5.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                            sqlcmdD5.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                            sqlcmdD5.Parameters.AddWithValue("@Lowest", double.Parse(lowest));
                                            sqlcmdD5.Parameters.AddWithValue("@LowestAdjusted", double.Parse(lowestadj));
                                            sqlcmdD5.Parameters.AddWithValue("@Opens", double.Parse(open));
                                            sqlcmdD5.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                            sqlcmdD5.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                            sqlcmdD5.ExecuteNonQuery();
                                        }
                                    }
                                    else
                                    {
                                        //Tiker, Date, Volume, OpenAdj, Open, LowestAdj, Lowest, HighestAdj
                                        string strQueryDetails6 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Open],[OpenAdjusted],[HighestAdjusted],[Lowest],[LowestAdjusted],[Volume]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted, @HighestAdjusted, @Lowest, @LowestAdjusted,@TotaTradingVolumes)";
                                        SqlCommand sqlcmdD6 = new SqlCommand(strQueryDetails6, dbcon);
                                        sqlcmdD6.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                        sqlcmdD6.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                        sqlcmdD6.Parameters.AddWithValue("@HighestAdjusted", double.Parse(highetsadj));
                                        sqlcmdD6.Parameters.AddWithValue("@Lowest", double.Parse(lowest));
                                        sqlcmdD6.Parameters.AddWithValue("@LowestAdjusted", double.Parse(lowestadj));
                                        sqlcmdD6.Parameters.AddWithValue("@Opens", double.Parse(open));
                                        sqlcmdD6.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                        sqlcmdD6.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                        sqlcmdD6.ExecuteNonQuery();
                                    }
                                }
                                else
                                {
                                    //Tiker, Date, Volume, OpenAdj, Open, LowestAdj, Lowest, HighestAdj, Highest
                                    string strQueryDetails7 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Open],[OpenAdjusted],[Highest],[HighestAdjusted],[Lowest],[LowestAdjusted],[Volume]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted, @Highest, @HighestAdjusted, @Lowest, @LowestAdjusted, @TotaTradingVolumes)";
                                    SqlCommand sqlcmdD7 = new SqlCommand(strQueryDetails7, dbcon);
                                    sqlcmdD7.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                    sqlcmdD7.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                    sqlcmdD7.Parameters.AddWithValue("@Highest", double.Parse(highet));
                                    sqlcmdD7.Parameters.AddWithValue("@HighestAdjusted", double.Parse(highetsadj));
                                    sqlcmdD7.Parameters.AddWithValue("@Lowest", double.Parse(lowest));
                                    sqlcmdD7.Parameters.AddWithValue("@LowestAdjusted", double.Parse(lowestadj));
                                    sqlcmdD7.Parameters.AddWithValue("@Opens", double.Parse(open));
                                    sqlcmdD7.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                    sqlcmdD7.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                    sqlcmdD7.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                //Tiker, Date, Volume, OpenAdj, Open, LowestAdj, Lowest, HighestAdj, Highest, CloseAdj
                                string strQueryDetails8 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Open],[OpenAdjusted],[Highest],[HighestAdjusted],[Lowest],[LowestAdjusted],[CloseAdjusted],[Volume]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted, @Highest, @HighestAdjusted, @Lowest, @LowestAdjusted, @CloseAdjusted, @TotaTradingVolumes)";
                                SqlCommand sqlcmdD8 = new SqlCommand(strQueryDetails8, dbcon);
                                sqlcmdD8.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                                sqlcmdD8.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                                sqlcmdD8.Parameters.AddWithValue("@CloseAdjusted", double.Parse(closeadj));
                                sqlcmdD8.Parameters.AddWithValue("@Highest", double.Parse(highet));
                                sqlcmdD8.Parameters.AddWithValue("@HighestAdjusted", double.Parse(highetsadj));
                                sqlcmdD8.Parameters.AddWithValue("@Lowest", double.Parse(lowest));
                                sqlcmdD8.Parameters.AddWithValue("@LowestAdjusted", double.Parse(lowestadj));
                                sqlcmdD8.Parameters.AddWithValue("@Opens", double.Parse(open));
                                sqlcmdD8.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                                sqlcmdD8.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                                sqlcmdD8.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            //Tiker, Date, Volume, OpenAdj, Open, LowestAdj, Lowest, HighestAdj, Highest, CloseAdj, Close
                            string strQueryDetails = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Open],[OpenAdjusted],[Highest],[HighestAdjusted],[Lowest],[LowestAdjusted],[Close],[CloseAdjusted],[Volume]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted, @Highest, @HighestAdjusted, @Lowest, @LowestAdjusted, @Closes, @CloseAdjusted, @TotaTradingVolumes)";
                            SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);
                            sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                            sqlcmdD.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                            sqlcmdD.Parameters.AddWithValue("@Closes", double.Parse(close));
                            sqlcmdD.Parameters.AddWithValue("@CloseAdjusted", double.Parse(closeadj));
                            sqlcmdD.Parameters.AddWithValue("@Highest", double.Parse(highet));
                            sqlcmdD.Parameters.AddWithValue("@HighestAdjusted", double.Parse(highetsadj));
                            sqlcmdD.Parameters.AddWithValue("@Lowest", double.Parse(lowest));
                            sqlcmdD.Parameters.AddWithValue("@LowestAdjusted", double.Parse(lowestadj));
                            sqlcmdD.Parameters.AddWithValue("@Opens", double.Parse(open));
                            sqlcmdD.Parameters.AddWithValue("@OpenAdjusted", double.Parse(openadj));
                            sqlcmdD.Parameters.AddWithValue("@TotaTradingVolumes", double.Parse(volume));
                            sqlcmdD.ExecuteNonQuery();
                        }
                    }
                }
                else
                {
                    for (int i = 8; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][0].ToString().Trim() == "")
                            break;
                        string closeadj = dt.Rows[i][4].ToString().Trim();
                        string close = dt.Rows[i][3].ToString().Trim();
                        if (closeadj == "")
                        {
                            if(close == "")
                            {
                                string strQueryDetailsNotClose1 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ) VALUES(@ticker,@trading)";
                                SqlCommand sqlcmdDNotClose1 = new SqlCommand(strQueryDetailsNotClose1, dbcon);
                                sqlcmdDNotClose1.Parameters.AddWithValue("@ticker", dt.Rows[i][0].ToString().Trim());
                                sqlcmdDNotClose1.Parameters.AddWithValue("@trading", dt.Rows[i][2].ToString().Trim());
                                sqlcmdDNotClose1.ExecuteNonQuery();
                            }
                            else
                            { 
                                string strQueryDetailsNotClose2 = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Close]) VALUES(@ticker,@trading, @Closes)";
                                SqlCommand sqlcmdDNotClose2 = new SqlCommand(strQueryDetailsNotClose2, dbcon);
                                sqlcmdDNotClose2.Parameters.AddWithValue("@ticker", dt.Rows[i][0].ToString().Trim());
                                sqlcmdDNotClose2.Parameters.AddWithValue("@trading", dt.Rows[i][2].ToString().Trim());
                                sqlcmdDNotClose2.Parameters.AddWithValue("@Closes", double.Parse(close));
                                sqlcmdDNotClose2.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            string strQueryDetails = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Close],[CloseAdjusted]) VALUES(@ticker,@trading, @Closes, @CloseAdjusted)";
                            SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);
                            sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][0].ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@trading", dt.Rows[i][2].ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@Closes", double.Parse(close));
                            sqlcmdD.Parameters.AddWithValue("@CloseAdjusted", double.Parse(closeadj));
                            sqlcmdD.ExecuteNonQuery();
                        }
                    }
                }
                dbcon.Close();
            }
        }
        private void SaveToDBIncome(System.Data.DataTable dt)
        {
            //string[] arr = new string[3];
            //arr[0] = "Year:";
            //arr[1] = "Quarter:";
            //arr[2] = "Unit:";
            //string[] arrOut;
            // List<string> listIDFeild = new List<string>();
            int dong = 0, cot = 0;
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();



                //// Save Feild
                //for (int i = 4; i <= dt.Columns.Count - 1; i++)
                //{
                //    string strQueryFeild = "IF NOT EXISTS (SELECT 1 FROM DBO.Income_Feild WHERE IDfeild=@idfeild ) BEGIN INSERT INTO [dbo].[Income_Feild] ([IDFeild],[unit]) VALUES (@idfeild,@unit) END";
                //    SqlCommand sqlcmd = new SqlCommand(strQueryFeild, dbcon);
                //    string pattern = dt.Rows[6][i].ToString();
                //    if (pattern.Trim() == "")
                //        break;
                //    //  string[] st = pattern.Split(new string[] { "Consolidated Year:", "Quarter:", "Unit:" }, StringSplitOptions.RemoveEmptyEntries);
                //    // print array st

                //    // Get Name
                //    int startPositionName = pattern.IndexOf(".") + ".".Length;
                //    string patternT = pattern.Substring(startPositionName, pattern.Length - startPositionName);

                //    int startPositionNameDot = patternT.IndexOf(".") + ".".Length;
                //    string patternTDot = patternT.Substring(startPositionNameDot, patternT.Length - startPositionNameDot);

                //    string name = patternTDot.Substring(0, patternTDot.IndexOf("Consolidated\nYear:"));
                //    // Get Unit
                //    int startUnitPosition = patternTDot.IndexOf("Unit:") + "Unit:".Length;
                //    string unit = patternTDot.Substring(startUnitPosition, patternTDot.Length - startUnitPosition);

                //    sqlcmd.Parameters.AddWithValue("@idfeild", name.Trim());
                //    sqlcmd.Parameters.AddWithValue("@unit", unit.Trim());

                //    sqlcmd.ExecuteNonQuery();
                //    cot++;
                //}




                //// Save Income_Financial
                //for (int i = 8; i <= dt.Rows.Count - 1; i++)
                //{
                //    try
                //    {
                //        if (dt.Rows[i][1].ToString().Trim() == "")
                //            break;
                //        // Save to Finance
                //        string strQueryFinance = "IF NOT EXISTS (SELECT 1 FROM DBO.Income_Financial WHERE Ticker=@Ticker ) BEGIN INSERT INTO [dbo].[Income_Financial] ([Ticker],[Name],[Exchange]) VALUES (@ticker,@name,@exchange) END";
                //        SqlCommand sqlcmd = new SqlCommand(strQueryFinance, dbcon);
                //        sqlcmd.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                //        sqlcmd.Parameters.AddWithValue("@name", dt.Rows[i][2].ToString().Trim());
                //        sqlcmd.Parameters.AddWithValue("@exchange", dt.Rows[i][3].ToString().Trim());
                //        sqlcmd.ExecuteNonQuery();
                //        dong++;
                //    }
                //    catch
                //    {
                //        //listBox2.Items.Add("SaveIncomeFinancial: " + i.ToString().Trim());
                //    }
                //}



                // demIncome += dong * cot;
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    try
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        for (int j = 4; j <= dt.Columns.Count - 1; j++)
                        {

                            //  string strQueryDetails = "INSERT INTO [dbo].[Income_Details_Feild]([Ticker] ,[IDFeild] ,[Year],[IDQuarter],[Value]) VALUES(@ticker,@explore,@year,@quarter,@value)";
                            string strq = "INSERT INTO [dbo].[Income] ([Ticker],[Year] ,[Quater]  ,[Name] ,[Value] ,[Unit]) VALUES(@ticker,@year,@quarter,@feildname,@value,@unit)";
                            SqlCommand sqlcmdD = new SqlCommand(strq, dbcon);

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

                            //Get unit
                            int startPositionUnit = patternTDot.IndexOf("Unit") + "Unit:".Length;
                            string unit = patternTDot.Substring(startPositionUnit, patternTDot.Length - startPositionUnit);
                            //Annual

                            sqlcmdD.Parameters.AddWithValue("@feildname", name.ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(year));
                            sqlcmdD.Parameters.AddWithValue("@quarter", quarter.ToString().Trim() != "Annual" ? Convert.ToInt16(quarter.ToString().Trim()) : 5);
                            sqlcmdD.Parameters.AddWithValue("@value", dt.Rows[i][j].ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@unit", unit.ToString().Trim());
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
        //Save to DB CashFlow
        private void SaveToDBCashFlow(System.Data.DataTable dt)
        {
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    try
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        for (int j = 4; j <= dt.Columns.Count - 1; j++)
                        {
                            string strq = "INSERT INTO [dbo].[CashFlow] ([Ticker],[Year] ,[Quater]  ,[Name] ,[Value] ,[Unit]) VALUES(@ticker,@year,@quarter,@feildname,@value,@unit)";
                            SqlCommand sqlcmdD = new SqlCommand(strq, dbcon);

                            sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                            string pattern = dt.Rows[6][j].ToString();
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

                            //Get unit
                            int startPositionUnit = patternTDot.IndexOf("Unit") + "Unit:".Length;
                            string unit = patternTDot.Substring(startPositionUnit, patternTDot.Length - startPositionUnit);
                            //Annual

                            sqlcmdD.Parameters.AddWithValue("@feildname", name.ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(year));
                            sqlcmdD.Parameters.AddWithValue("@quarter", quarter.ToString().Trim() != "Annual" ? Convert.ToInt16(quarter.ToString().Trim()) : 5);
                            sqlcmdD.Parameters.AddWithValue("@value", dt.Rows[i][j].ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@unit", unit.ToString().Trim());
                            sqlcmdD.ExecuteNonQuery();
                            if (dt.Rows[i][1].ToString().Trim() == "")
                                break;
                        }
                    }
                    catch
                    {
                    }
                }
                dbcon.Close();
            }

        }
        //Note
        private void SaveToDBNote(System.Data.DataTable dt)
        {
            using (SqlConnection dbcon = new SqlConnection(sqlConnectionString))
            {
                dbcon.Open();
                for (int i = 8; i <= dt.Rows.Count - 1; i++)
                {
                    try
                    {
                        if (dt.Rows[i][1].ToString().Trim() == "")
                            break;
                        for (int j = 4; j <= dt.Columns.Count - 1; j++)
                        {
                            string strq = "INSERT INTO [dbo].[Note] ([Ticker],[Year] ,[Quater]  ,[Name] ,[Value] ,[Unit]) VALUES(@ticker,@year,@quarter,@feildname,@value,@unit)";
                            SqlCommand sqlcmdD = new SqlCommand(strq, dbcon);

                            sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                            string pattern = dt.Rows[6][j].ToString();
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

                            //Get unit
                            int startPositionUnit = patternTDot.IndexOf("Unit") + "Unit:".Length;
                            string unit = patternTDot.Substring(startPositionUnit, patternTDot.Length - startPositionUnit);
                            //Annual

                            sqlcmdD.Parameters.AddWithValue("@feildname", name.ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(year));
                            sqlcmdD.Parameters.AddWithValue("@quarter", quarter.ToString().Trim() != "Annual" ? Convert.ToInt16(quarter.ToString().Trim()) : 5);
                            sqlcmdD.Parameters.AddWithValue("@value", dt.Rows[i][j].ToString().Trim());
                            sqlcmdD.Parameters.AddWithValue("@unit", unit.ToString().Trim());
                            sqlcmdD.ExecuteNonQuery();
                            if (dt.Rows[i][1].ToString().Trim() == "")
                                break;
                        }
                    }
                    catch
                    {
                    }
                }
                dbcon.Close();
            }

        }
        #endregion

        #region Store file
        //Income
        private void StoreFileIncome(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);
                if (!Directory.Exists(tbOutput.Text + "\\Income\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\Income\\");
                }
                File.Copy(fileName, tbOutput.Text + "\\Income\\" + filenameOnly, true);
                File.Delete(fileName);
            }
        }
        //Cashflow
        private void StoreFileCashFlow(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);
                if (!Directory.Exists(tbOutput.Text + "\\cashflow\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\cashflow\\");
                }
                File.Copy(fileName, tbOutput.Text + "\\cashflow\\" + filenameOnly, true);
                File.Delete(fileName);
            }
        }
        //Note
        private void StoreFileNote(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);
                if (!Directory.Exists(tbOutput.Text + "\\Note\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\Note\\");
                }
                File.Copy(fileName, tbOutput.Text + "\\Note\\" + filenameOnly, true);
                File.Delete(fileName);
            }
        }
        //Ratios
        private void StoreFileRatios(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);
                if (!Directory.Exists(tbOutput.Text + "\\Ratios\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\Ratios\\");
                }
                File.Copy(fileName, tbOutput.Text + "\\Ratios\\" + filenameOnly, true);
                File.Delete(fileName);
            }
        }
        //Balance
        private void StoreFileBalance(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);
                if (!Directory.Exists(tbOutput.Text + "\\Balance\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\Balance\\");
                }
                File.Copy(fileName, tbOutput.Text + "\\Balance\\" + filenameOnly, true);
                File.Delete(fileName);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void bntCashFlow_Click(object sender, EventArgs e)
        {
            bntCashFlow.Text = "CashFlow Running...";
            MoveToTempCashFlow();
            ReadExcelAndSaveCashFlow();
            bntCashFlow.Text = "CashFlow";
        }

        //Stock Market Data
        private void StoreFileStock(string fileName)
        {
            if (File.Exists(fileName))
            {
                string filenameOnly = Path.GetFileName(fileName);
                if (!Directory.Exists(tbOutput.Text + "\\Stock\\"))
                {
                    Directory.CreateDirectory(tbOutput.Text + "\\Stock\\");
                }
                File.Copy(fileName, tbOutput.Text + "\\Stock\\" + filenameOnly, true);
                File.Delete(fileName);
            }
        }

        private void tbInput_TextChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void btnNote_Click(object sender, EventArgs e)
        {
            btnNote.Text = "Note Running...";
            MoveToTempNote();
            ReadExcelAndSaveNote();
            btnNote.Text = "Note";
        }
        #endregion

        #region Read Excel And Save
        //Income
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

                        string fullNameIn_In_Out = tbOutput.Text + "Income\\" + fileName.Replace(".", string.Empty) + ".xls";
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
        //Ratios
        void ReadExcelAndSaveRatios()
        {
            try
            {
                var files1 = Directory.GetFiles(pathTempRatios, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls"));
                foreach (string file in files1)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        txtStatus.Text = fileName + " is loading to Server";
                        string fullNameIn_In_Out = tbOutput.Text + "Ratios\\" + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {
                            txtStatus.Text = (fileName + " is loading to Server").ToString();
                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBRatio(tbl);
                                //progressBar1.Value++;
                                break;
                            }

                        }
                        StoreFileRatios(file);
                        txtStatus.Text = "";
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
        //Balance
        void ReadExcelAndSaveBalance()
        {
            try
            {
                var files1 = Directory.GetFiles(pathTempBalance, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls"));
                foreach (string file in files1)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        string fullNameIn_In_Out = tbOutput.Text + "Balance\\" + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {
                            txtStatus.Text = (fileName + " is loading to Server").ToString();
                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBBalance(tbl);
                                break;
                            }

                        }
                        StoreFileBalance(file);
                        txtStatus.Text = "";
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
        //Stock Market Data
        void ReadExcelAndSaveStock()
        {
            try
            {
                var files1 = Directory.GetFiles(pathTempStock, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls"));
                foreach (string file in files1)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        txtStatus.Text = fileName + " is loading to Server";
                        string fullNameIn_In_Out = tbOutput.Text + "Stock\\" + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {
                            txtStatus.Text = (fileName + " is loading to Server").ToString();
                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBStock(tbl);
                                //progressBar1.Value++;
                                break;
                            }

                        }
                        StoreFileStock(file);
                        txtStatus.Text = "";
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
        //CashFlow
        void ReadExcelAndSaveCashFlow()
        {
            try
            {
                var files1 = Directory.GetFiles(pathTempCashFlow, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls"));
                foreach (string file in files1)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);

                        string fullNameIn_In_Out = tbOutput.Text + "Cashflow\\" + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {

                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBCashFlow(tbl);
                                break;
                            }

                        }
                        StoreFileCashFlow(file);
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
        void ReadExcelAndSaveNote()
        {
            try
            {
                var files1 = Directory.GetFiles(pathTempNote, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls"));
                foreach (string file in files1)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);

                        string fullNameIn_In_Out = tbOutput.Text + "note\\" + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {

                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBNote(tbl);
                                break;
                            }

                        }
                        StoreFileNote(file);
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
        #endregion
    }

}



