using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RecursiveFileSearch
{
    class Program
    {
        static List<string> allFiles = new List<string>();
        static string fileSearchPath = "";
        //[STAThread]
        static void Main(string[] args)
        {
            KillExcelProcess();

            //FolderBrowserDialog fbd = new FolderBrowserDialog();
            //if (fbd.ShowDialog() == DialogResult.OK)
            //{
            //    fileSearchPath = fbd.SelectedPath;
            //}
            //else
            //{
                fileSearchPath = ConfigurationManager.AppSettings["FolderPath"].ToString();
           // }

            // Get all Files including Paths
            foreach (string file in Directory.EnumerateFiles(fileSearchPath, "*.*", SearchOption.AllDirectories))
            {
                FileInfo fI = new FileInfo(file);

                string[] extensions = ConfigurationManager.AppSettings["FilesToIncludeInSearch"].ToString().Split(',');

                foreach (string ext in extensions)
                {
                    if (fI.Extension.ToLower().Equals(ext.ToLower()) || ext.Equals("*.*"))
                    {
                        allFiles.Add(file);
                    }
                }
            }

            WriteFilesToExcel();
            KillExcelProcess();
        }

        private static void WriteFilesToExcel()
        {
            Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Excel._Workbook excelWorkBook;
            Microsoft.Office.Interop.Excel._Worksheet excelWorkSheet;
            Microsoft.Office.Interop.Excel.Range worksheetRange;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;

                //Get a new workbook.
                excelWorkBook = (Microsoft.Office.Interop.Excel._Workbook)(excelApp.Workbooks.Add(""));
                excelWorkSheet = (Microsoft.Office.Interop.Excel._Worksheet)excelWorkBook.ActiveSheet;
                excelWorkSheet.Name = ConfigurationManager.AppSettings["ExcelReportSheetName"].ToString();
                excelWorkSheet.Columns.AutoFit();
                worksheetRange = excelWorkSheet.Range["A1", "D" + allFiles.Count];

                // Set the Width of Excel Columns
                int ColumnFileNameWidth = 25;
                int ColumnFilePathWidth = 25;

                foreach (string filePath in allFiles)
                {
                    FileInfo fileI = new FileInfo(filePath);
                    if (fileI.Name.Length > ColumnFileNameWidth)
                    {
                        ColumnFileNameWidth = fileI.Name.Length;
                    }

                    if (filePath.Length > ColumnFilePathWidth)
                    {
                        ColumnFilePathWidth = filePath.Length;
                    }
                }

                worksheetRange.Columns[1].ColumnWidth = ColumnFileNameWidth;
                worksheetRange.Columns[2].ColumnWidth = 10;
                worksheetRange.Columns[3].ColumnWidth = ColumnFilePathWidth;
                worksheetRange.Columns[4].ColumnWidth = 15;
                worksheetRange.Columns[5].ColumnWidth = 10;

                // File Headers
                worksheetRange[1, 1] = "FILE NAME";
                worksheetRange[1, 1].EntireRow.Font.Bold = true;
                worksheetRange[1, 1].Interior.Color = XlRgbColor.rgbLightSkyBlue;
                worksheetRange[1, 2] = "FILE TYPE";
                worksheetRange[1, 2].EntireRow.Font.Bold = true;
                worksheetRange[1, 2].Interior.Color = XlRgbColor.rgbLightSkyBlue;
                worksheetRange[1, 3] = "FILE PATH";
                worksheetRange[1, 3].EntireRow.Font.Bold = true;
                worksheetRange[1, 3].Interior.Color = XlRgbColor.rgbLightSkyBlue;
                worksheetRange[1, 4] = "DATE MODIFIED";
                worksheetRange[1, 4].EntireRow.Font.Bold = true;
                worksheetRange[1, 4].Interior.Color = XlRgbColor.rgbLightSkyBlue;
                worksheetRange[1, 4].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                worksheetRange[1, 5] = "SIZE";
                worksheetRange[1, 5].EntireRow.Font.Bold = true;
                worksheetRange[1, 5].Interior.Color = XlRgbColor.rgbLightSkyBlue;

                for (int k = 0; k < allFiles.Count; k++) // k= Excel Row
                {
                    FileInfo fileInfo = new FileInfo(allFiles[k]);
                    for (int i = 1; i <= 5; i++) // i = Excel Column
                    {
                        if (i == 1)
                        {
                            worksheetRange[k+2, i] = fileInfo.Name;
                        }
                        if (i == 2)
                        {
                            worksheetRange[k+2, i] = fileInfo.Extension;
                        }
                        if (i == 3)
                        {
                            worksheetRange[k+2, i] = (allFiles[k].Replace(fileSearchPath, ConfigurationManager.AppSettings["Host"].ToString())).Replace('\\', '/');
                        }
                        if (i == 4)
                            worksheetRange[k+2, i] = fileInfo.LastWriteTime.ToString(ConfigurationManager.AppSettings["DateFormat"].ToString());
                        if (i == 5)
                        {
                            worksheetRange[k+2, i] = (fileInfo.Length < 1024) ? fileInfo.Length.ToString() + " B" : (fileInfo.Length / 1024) + " KB";
                        }
                    }
                }

                excelWorkBook.SaveAs(string.Format(ConfigurationManager.AppSettings["ExcelReportFileName"].ToString(),DateTime.Now.ToString("yyyy-MM-dd-HHmmss")));
                excelWorkBook.Close();

                #region COM CLEANUP
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(worksheetRange);
                Marshal.FinalReleaseComObject(excelWorkSheet);

                Marshal.FinalReleaseComObject(excelWorkBook);

                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
                #endregion 
            }
            catch (Exception ex)
            {
                using (StreamWriter logWriter = new StreamWriter(ConfigurationManager.AppSettings["LogFileName"].ToString(), true))
                {
                    logWriter.Write("ERROR IN APPLICATION" + ex.Message + ex.StackTrace);
                }
            }

        }

        private static void KillExcelProcess()
        {
            try
            {
                var processes = from p in Process.GetProcessesByName("EXCEL")
                                select p;

                foreach (var process in processes)
                {
                    process.Kill();                    
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter logWriter = new StreamWriter(ConfigurationManager.AppSettings["LogFileName"].ToString(), true))
                {
                    logWriter.Write("ERROR IN APPLICATION" + ex.Message + ex.StackTrace);
                }
            }
        }
    }
}
