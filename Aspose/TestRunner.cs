using Aspose.Cells;
using System.Diagnostics;
using System.Globalization;

namespace Aspose
{
    public class TestRunner
    {
        public static TimeSpan[] RunTests()
        {
            var timeTable = new TimeSpan[10];

            timeTable[0] = Run320000RandomCellsTest();
            timeTable[1] = Run160000DateCellsTest();
            timeTable[2] = RunStyleChangesTest();
            timeTable[3] = GetTimeSpan();
            timeTable[4] = GetTimeSpan();
            timeTable[5] = GetTimeSpan();
            timeTable[6] = GetTimeSpan();
            timeTable[7] = GetTimeSpan();
            timeTable[8] = GetTimeSpan();
            timeTable[9] = GetTimeSpan();

            return timeTable;
        }

        private static TimeSpan GetTimeSpan()
        {
            return TimeSpan.FromSeconds(10);
        }

        private static string GetRandomDate(Random gen)
        {
            DateTime start = new(1995, 1, 1);
            int range = (DateTime.Today - start).Days;
            return start.AddDays(gen.Next(range)).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        private static decimal GetRandomDecimal(Random rng)
        {
            byte scale = (byte)rng.Next(29);
            bool sign = rng.Next(2) == 1;
            return new decimal(GetRandomRandInt(rng),
                GetRandomRandInt(rng),
                GetRandomRandInt(rng),
                sign,
                scale);
        }

        private static int GetRandomRandInt(Random rng)
        {
            int firstBits = rng.Next(0, 1 << 4) << 28;
            int lastBits = rng.Next(0, 1 << 28);
            return firstBits | lastBits;
        }

        private static TimeSpan Run320000RandomCellsTest()
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

        private static TimeSpan Run160000DateCellsTest()
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

        private static TimeSpan RunStyleChangesTest()
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