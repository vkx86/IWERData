using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace IWERDataCollector
{
    internal class WorkbooksDataCollectorError : Exception
    {
        internal WorkbooksDataCollectorError(string message) : base(message) { }
    }

    internal class WorkbooksDataCollector
    {
        private readonly string _workbooksDirectory;
        private ExcelApp._Application _excelApp;

        internal WorkbooksDataCollector(string workbooksDirectory)
        {
            _workbooksDirectory = workbooksDirectory;
        }

        internal void CollectData(IEnumerable<string> years)
        {
            _excelApp = new ExcelApp.Application();

            foreach (var year in years)
                CollectDataForYear(year);

            _excelApp.Quit();
        }

        private void CollectDataForYear(string year)
        {
            var workbooks = EnumerateWorkbooks(year).ToList();
            if ( !workbooks.Any())
                throw new WorkbooksDataCollectorError($"No ER workbooks for year {year} found!");

            using (var writer = new StreamWriter(Path.Combine(_workbooksDirectory, $"ER_{year}.csv")))
            {
                writer.WriteLine(
                    "week,total,total age < 1,pneumonia,pneumonia age < 1,pneumonia + other,pneumonia + other age < 1");
                foreach (var filename in workbooks)
                {
                    Console.WriteLine($"{Path.GetFileName(filename)}:");
                    CollectDataFromWorkbook(filename, writer);
                    Console.WriteLine("");
                }
            }
        }

        private IEnumerable<string> EnumerateWorkbooks(string year)
        {
            if ( !Directory.Exists(_workbooksDirectory))
                throw new WorkbooksDataCollectorError("Workbooks directory not found!");

            return Directory.EnumerateFiles(_workbooksDirectory, "*.xlsx")
                .Where(file => !Path.GetFileNameWithoutExtension(file).Contains("~") && Path.GetFileNameWithoutExtension(file).Contains(year))
                .ToList();
        }

        private void CollectDataFromWorkbook(string filename, TextWriter writer)
        {
            var workbook = _excelApp.Workbooks.Open(filename);
            new DeathDataCollector(workbook, writer).Collect();
            workbook.Close();
        }

    }

}
