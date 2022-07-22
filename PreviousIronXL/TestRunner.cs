using IronXLOld;
using IronXLOld.Styles;
using System.Diagnostics;

namespace PreviousIronXL
{
    public class TestRunner : TestRunnerBase.TestRunner
    {
        public override TimeSpan Run320000RandomCellsTest()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var worksheet = workbook.DefaultWorkSheet;

            var rand = new Random();
            for (int i = 1; i <= 20000; i++)
            {
                worksheet["A" + i].Value = $"=\"{Guid.NewGuid()}\"";
                worksheet["B" + i].Value = $"=\"{Guid.NewGuid()}\"";
                worksheet["C" + i].Value = Guid.NewGuid().ToString();
                worksheet["D" + i].Value = rand.Next(32);
                worksheet["E" + i].Value = $"=\"{Guid.NewGuid()}\"";
                worksheet["F" + i].Value = $"=\"{Guid.NewGuid()}\"";
                worksheet["G" + i].Value = Guid.NewGuid().ToString();
                worksheet["H" + i].Value = rand.Next(13);
                worksheet["I" + i].Value = GetRandomDate(rand);
                worksheet["J" + i].Value = GetRandomDate(rand);
                worksheet["K" + i].Value = Guid.NewGuid().ToString();
                worksheet["L" + i].Value = $"=\"{Guid.NewGuid()}\"";
                worksheet["M" + i].Value = Guid.NewGuid().ToString();
                worksheet["N" + i].Value = Guid.NewGuid().ToString();
                worksheet["O" + i].Value = GetRandomDecimal(rand);
                worksheet["P" + i].Value = GetRandomDecimal(rand);
            }

            workbook.SaveAs("PreviousIXLRandomCells.xlsx");

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan Run160000DateCellsTest()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var worksheet = workbook.DefaultWorkSheet;

            int rowNo = 80000;
            
            for (int i = 1; i < rowNo; i++)
            {
                worksheet["A" + i].Value = i + 1;
                worksheet["B" + i].Value = DateTime.Now;
            }

            workbook.SaveAs("PreviousIXLDateCells.xlsx");

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan RunStyleChangesTest()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var worksheet = workbook.DefaultWorkSheet;

            worksheet.InsertRows(19, 319);

            var range = worksheet.GetRange("I7:O319");
            range.Value = "Value";
            
            var style = range.Style;
            
            style.Font.Height = 22;
            style.VerticalAlignment = VerticalAlignment.Bottom;
            style.HorizontalAlignment = HorizontalAlignment.Left;

            workbook.SaveAs("PreviousIXLStyleChange.xlsx");

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}