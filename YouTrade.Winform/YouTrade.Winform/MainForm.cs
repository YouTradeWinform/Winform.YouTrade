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
        public int demFileRatios =0,demFileBalance=0,demFileStock=0;
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
            if(!Directory.Exists(pathTempCashFlow))
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
            btnRatios.Text = "Ratios Running...";
            btnRatios.BackColor = Color.Blue;
            MoveToTempRatios();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 1;
            progressBar1.Step = 10;
            for (int i = 0; i < demFileRatios; i++)
            {
                ReadExcelAndSaveRatios();
            }
            txtStatus.Text = "Done!";
            btnRatios.Text = "Ratios";
        }
        private void Click_Balance(object sender, EventArgs e)
        {
            btnBalance.Text = "Balance Running...";
            btnRatios.BackColor = Color.Blue;
            MoveToTempBalance();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 1;
            progressBar1.Step = 10;
            for (int i = 0; i < demFileBalance; i++)
            {
                ReadExcelAndSaveBalance();
            }
            txtStatus.Text = "Done!";
            btnBalance.Text = "Balance";
        }
        private void Click_Stock(object sender, EventArgs e)
        {
            btnStock.Text = "Stock Running...";
            btnRatios.BackColor = Color.Blue;
            MoveToTempStock();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 1;
            progressBar1.Step = 10;
            for (int i = 0; i < demFileStock; i++)
            {
                ReadExcelAndSaveStock();
            }
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
                        //Unit
                        int startUnitPosition = pattern.IndexOf("Unit:") + "Unit:".Length;
                        string unit = pattern.Substring(startUnitPosition, pattern.Length - startUnitPosition);
                        if (true == (unit.Contains("\n")))
                            unit = unit.Substring(0, unit.Length);

                        sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@explorename", name.ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(years));
                        sqlcmdD.Parameters.AddWithValue("@quater", quarter.ToString().Trim() != "Annual" ? Convert.ToInt16(quarter.ToString().Trim()) : 5);
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
                        string strQueryDetails = "INSERT INTO [dbo].[MarketData]([Ticker] ,[Date] ,[Opens],[OpenAdjusted],[Highest],[HighestAdjusted],[Lowest],[LowestAdjusted],[Closes],[CloseAdjusted],[TotaTradingVolumes]) VALUES(@ticker,@trading, @Opens, @OpenAdjusted, @Highest, @HighestAdjusted, @Lowest, @LowestAdjusted, @Closes, @CloseAdjusted, @TotaTradingVolumes)";
                        SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);
                        sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                        sqlcmdD.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                        sqlcmdD.Parameters.AddWithValue("@Closes", dt.Rows[i][5]);
                        sqlcmdD.Parameters.AddWithValue("@CloseAdjusted", dt.Rows[i][6]);
                        sqlcmdD.Parameters.AddWithValue("@Highest", dt.Rows[i][22]);
                        sqlcmdD.Parameters.AddWithValue("@HighestAdjusted", dt.Rows[i][23]);
                        sqlcmdD.Parameters.AddWithValue("@Lowest", dt.Rows[i][24]);
                        sqlcmdD.Parameters.AddWithValue("@LowestAdjusted", dt.Rows[i][25]);
                        sqlcmdD.Parameters.AddWithValue("@Opens", dt.Rows[i][26]);
                        sqlcmdD.Parameters.AddWithValue("@OpenAdjusted", dt.Rows[i][27]);
                        sqlcmdD.Parameters.AddWithValue("@TotaTradingVolumes", dt.Rows[i][40]);
                        sqlcmdD.ExecuteNonQuery();
                    }
                }
                else
                {
                    for (int i = 8; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][0].ToString().Trim() == "")
                            break;
                        string strQueryDetails = "INSERT INTO [dbo].[StockMarketData]([Ticker] ,[Trading] ,[Closes],[CloseAdjusted]) VALUES(@ticker,@trading, @Closes, @CloseAdjusted)";
                        SqlCommand sqlcmdD = new SqlCommand(strQueryDetails, dbcon);
                        sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][0]);
                        sqlcmdD.Parameters.AddWithValue("@trading", dt.Rows[i][2]);
                        sqlcmdD.Parameters.AddWithValue("@Closes", dt.Rows[i][3]);
                        sqlcmdD.Parameters.AddWithValue("@CloseAdjusted", dt.Rows[i][4]);
                        sqlcmdD.ExecuteNonQuery();
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
                        //Unit
                        int startUnitPosition = pattern.IndexOf("Unit:") + "Unit:".Length;
                        string unit = pattern.Substring(startUnitPosition, pattern.Length - startUnitPosition);
                        if (true == (unit.Contains("\n")))
                            unit = unit.Substring(0, unit.Length);

                        sqlcmdD.Parameters.AddWithValue("@ticker", dt.Rows[i][1].ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@explorename", name.ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@year", Convert.ToInt16(years));
                        sqlcmdD.Parameters.AddWithValue("@quater", quarter.ToString().Trim() != "Annual" ? Convert.ToInt16(quarter.ToString().Trim()) : 5);
                        sqlcmdD.Parameters.AddWithValue("@value", dt.Rows[i][j].ToString().Trim());
                        sqlcmdD.Parameters.AddWithValue("@unit", unit.ToString().Trim());
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
                            sqlcmdD.Parameters.AddWithValue("@quarter", quarter.ToString().Trim() != "Annual" ? Convert.ToInt16(quarter.ToString().Trim()) : 5 );
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

                        string fullNameIn_In_Out = tbOutput.Text + "Ratios\\"+ fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {

                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                txtStatus.Text = fileName;
                                SaveToDBRatio(tbl);
                                break;
                            }

                        }
                        StoreFileRatios(file);
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
                            txtStatus.Text = fileName;
                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBBalance(tbl);
                                break;
                            }

                        }
                        StoreFileBalance(file);
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

                        string fullNameIn_In_Out = tbOutput.Text + "Stock\\" + fileName.Replace(".", string.Empty) + ".xls";
                        if (!File.Exists(fullNameIn_In_Out))
                        {
                            txtStatus.Text = fileName;
                            dsSource = GetDatasetFromExcel(file);
                            foreach (System.Data.DataTable tbl in dsSource.Tables)
                            {
                                SaveToDBStock(tbl);
                                break;
                            }

                        }
                        StoreFileStock(file);
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



