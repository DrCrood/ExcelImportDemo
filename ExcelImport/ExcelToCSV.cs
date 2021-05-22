using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Threading.Tasks;
using System.IO;

namespace LLPAPricingExcelImport
{
    class ExcelToCSV
    {
        public string file { get; set; }
        public List<string> CSVTextBody { get; set; }

        public ExcelToCSV(string filename)
        {
            file = filename;
            CSVTextBody = new List<string>();
        }
        public void ConvertExcelToCSV()
        {

            List<string> csvRows = new List<string>();

            using XLWorkbook workbook = new XLWorkbook(file);

            IXLWorksheet worksheet = workbook.Worksheet(1);

            foreach (IXLRow row in worksheet.RowsUsed())
            {
                csvRows.Clear();
                foreach (IXLCell cell in row.Cells(false))
                {
                    var v2 = cell.Value.ToString();

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
                CSVTextBody.Add(newLine);
                //Console.WriteLine(CSVTextBody.Count);
            }

            File.WriteAllLines(file.Substring(0, file.Length - 4) + "CSV", CSVTextBody);
        }
    }
}
