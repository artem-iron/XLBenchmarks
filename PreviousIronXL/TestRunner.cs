using IronXLOld;
using IronXLOld.Styles;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;

namespace PreviousIronXL
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

        private static TimeSpan Run160000DateCellsTest()
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

        private static TimeSpan RunStyleChangesTest()
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