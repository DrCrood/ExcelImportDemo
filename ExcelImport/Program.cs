using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ClosedXML;

namespace LLPAPricingExcelImport
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();
            string file = @"path-to-the-file\test.xlsx";

            //new ExcelToCSV(file).ConvertExcelToCSV();

            new ExcelToCSV_Microsoft(file).ConvertExcelToCSV();

            watch.Stop();
            Console.WriteLine(watch.Elapsed.TotalSeconds);
        }
    }
}
