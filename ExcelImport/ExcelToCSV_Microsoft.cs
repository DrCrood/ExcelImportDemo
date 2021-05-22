using System;
using System.Collections.Generic;
using System.Runtime;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;

namespace LLPAPricingExcelImport
{
    class ExcelToCSV_Microsoft
    {
        public string file { get; set; }

        public ExcelToCSV_Microsoft(string file)
        {
            this.file = file;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out Int32 lpdwProcessId);

        public void ConvertExcelToCSV()
        {
            Application exApp = new Application();
            int n = exApp.Application.Hwnd;
            Workbook rateBook = exApp.Workbooks.Open(file);
            var rateSheet = (Worksheet)rateBook.ActiveSheet;

            Microsoft.Office.Interop.Excel.Range excelRange = rateSheet.UsedRange;

            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            object[,] values = excelRange.Value2;

            List<string> csvRows = new List<string>();
            List<string> csvBody = new List<string>();
            for (int i = 1; i <= rowCount; i++)
            {
                csvRows.Clear();
                for (int j = 1; j <= colCount; j++)
                {
                    //var cell = excelRange.Cells[i, j];
                    //var v2 = cell?.Value2 as String;
                    var cell = values[i, j];
                    var v2 = cell.ToString();

                    if (cell == null || v2 == null)
                    {
                        csvRows.Add(" ");
                    }
                    else
                    {
                        csvRows.Add(v2);
                    }
                }
                var newLine = string.Join(",", csvRows);
                csvBody.Add(newLine);
                //Console.WriteLine(csvBody.Count);
            }

            File.WriteAllLines(file.Substring(0,file.Length-4)+"CSV", csvBody);
            GetWindowThreadProcessId((IntPtr)n, out int pid);
            System.Diagnostics.Process.GetProcessById(pid).Kill();
        }
    }

}
