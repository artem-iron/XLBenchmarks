using Aspose.Cells;
using System.Diagnostics;

namespace Aspose
{
    public class TestRunner : TestRunnerBase.TestRunner
    {
        public override TimeSpan Run320000RandomCellsTest()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            var rand = new Random();
            for (int i = 1; i <= 20000; i++)
            {
                cells["A" + i].Value = $"=\"{Guid.NewGuid()}\"";
                cells["B" + i].Value = $"=\"{Guid.NewGuid()}\"";
                cells["C" + i].Value = Guid.NewGuid().ToString();
                cells["D" + i].Value = rand.Next(32);
                cells["E" + i].Value = $"=\"{Guid.NewGuid()}\"";
                cells["F" + i].Value = $"=\"{Guid.NewGuid()}\"";
                cells["G" + i].Value = Guid.NewGuid().ToString();
                cells["H" + i].Value = rand.Next(13);
                cells["I" + i].Value = GetRandomDate(rand);
                cells["J" + i].Value = GetRandomDate(rand);
                cells["K" + i].Value = Guid.NewGuid().ToString();
                cells["L" + i].Value = $"=\"{Guid.NewGuid()}\"";
                cells["M" + i].Value = Guid.NewGuid().ToString();
                cells["N" + i].Value = Guid.NewGuid().ToString();
                cells["O" + i].Value = GetRandomDecimal(rand);
                cells["P" + i].Value = GetRandomDecimal(rand);
            }

            workbook.Save("AsposeRandomCells.xlsx");

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan Run160000DateCellsTest()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            int rowNo = 80000;
            
            for (int i = 1; i < rowNo; i++)
            {
                cells["A" + i].Value = i + 1;
                cells["B" + i].Value = DateTime.Now;
            }

            workbook.Save("AsposeDateCells.xlsx");

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan RunStyleChangesTest()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            cells.InsertRows(19, 319);

            var range = cells.CreateRange("I7:O319");
            range.Value = "Value";
            
            var style = new CellsFactory().CreateStyle();

            style.Font.Size = 22;
            style.VerticalAlignment = TextAlignmentType.Bottom;
            style.HorizontalAlignment = TextAlignmentType.Left;

            range.ApplyStyle(style, new StyleFlag() { Font = true, VerticalAlignment = true, HorizontalAlignment = true });

            workbook.Save("AsposeStyleChange.xlsx");
            
            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}