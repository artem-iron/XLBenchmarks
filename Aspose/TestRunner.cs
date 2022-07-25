using Aspose.Cells;
using System.Diagnostics;

namespace Aspose
{
    public class TestRunner : TestRunnerBase.TestRunner
    {
        public override string TestRunnerName => typeof(TestRunner).Namespace ?? "Aspose";

        public override TimeSpan RunRandomCellsTest(bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            var rand = new Random();
            for (int i = 1; i <= RandomCellsRowNumber; i++)
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

            if (savingResultingFile)
            {
                workbook.Save(RandomCellsFileName);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan RunDateCellsTest(bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            for (int i = 1; i < DateCellsNumber; i++)
            {
                cells["A" + i].Value = DateTime.Now;
            }

            if (savingResultingFile)
            {
                workbook.Save(DateCellsFileName);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public override TimeSpan RunStyleChangesTest(bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            cells.InsertRows(1, StyleChangeRowNumber);

            var range = cells.CreateRange($"A1:O{StyleChangeRowNumber}");
            range.Value = CELL_VALUE;
            
            var style = new CellsFactory().CreateStyle();

            style.Font.Size = 22;
            style.VerticalAlignment = TextAlignmentType.Top;
            style.HorizontalAlignment = TextAlignmentType.Right;

            range.ApplyStyle(style, new StyleFlag() { Font = true, VerticalAlignment = true, HorizontalAlignment = true });

            if (savingResultingFile)
            {
                workbook.Save(StyleChangeFileName);
            }

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}