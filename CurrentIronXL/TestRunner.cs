using IronXL;
using IronXL.Styles;
using System.Diagnostics;

namespace CurrentIronXL
{
    public class TestRunner : TestRunnerBase.TestRunner
    {
        public override string TestRunnerName => typeof(TestRunner).Namespace ?? "CurrentIronXL";

        public override TimeSpan RunRandomCellsTest(bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var worksheet = workbook.DefaultWorkSheet;

            var rand = new Random();
            for (int i = 1; i <= RandomCellsRowNumber; i++)
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

            if (savingResultingFile)
            {
                workbook.SaveAs(RandomCellsFileName);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan RunDateCellsTest(bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var worksheet = workbook.DefaultWorkSheet;

            for (int i = 1; i < DateCellsNumber; i++)
            {
                worksheet["A" + i].Value = DateTime.Now;
            }

            if (savingResultingFile)
            {
                workbook.SaveAs(DateCellsFileName);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan RunStyleChangesTest(bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var worksheet = workbook.DefaultWorkSheet;

            worksheet.InsertRows(1, StyleChangeRowNumber);

            var range = worksheet.GetRange($"A1:O{StyleChangeRowNumber}");
            range.Value = CELL_VALUE;

            var style = range.Style;

            style.Font.Height = 22;
            style.VerticalAlignment = VerticalAlignment.Top;
            style.HorizontalAlignment = HorizontalAlignment.Right;

            if (savingResultingFile)
            {
                workbook.SaveAs(StyleChangeFileName);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}