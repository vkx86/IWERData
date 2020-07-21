using System;
using System.IO;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace IWERDataCollector
{
    internal class DeathDataCollector
    {
        private readonly ExcelApp.Range _range;
        private readonly TextWriter _writer;
        private const int ValuesCol = 17;

        internal DeathDataCollector(ExcelApp._Workbook workbook, TextWriter writer)
        {
            _range = (workbook.Sheets[1] as ExcelApp._Worksheet)?.UsedRange;
            _writer = writer;
        }

        internal void Collect()
        {
            var line = 
                $"{Week},{Total},{TotalAge1},{Pneumonia},{PneumoniaAge1},{PneumoniaOther},{PneumoniaOtherAge1}";

            _writer.WriteLine(line);
            Console.WriteLine(line);
        }

        private string Week =>
            ReadCell(14, 1);

        private string Total =>
            ReadCell(132, ValuesCol);

        private string TotalAge1 =>
            ReadCell(133, ValuesCol);

        private string Pneumonia =>
            ReadCell(134, ValuesCol);

        private string PneumoniaAge1 =>
            ReadCell(135, ValuesCol);

        private string PneumoniaOther =>
            ReadCell(136, ValuesCol);

        private string PneumoniaOtherAge1 =>
            ReadCell(137, ValuesCol);

        private string ReadCell(int row, int col)
        {
            return ((ExcelApp.Range)_range.Cells[row, col]).Value2.ToString();
        }

    }

}
