using IronXL;
using System.Diagnostics;
using System.Globalization;

namespace CurrentIronXL
{
    public class TestRunner
    {
        public static TimeSpan[] RunTests()
        {
            var timeTable = new TimeSpan[10];

            timeTable[0] = Run80000RandomCellsTest();
            timeTable[1] = Run80000RandomCellsTest();
            timeTable[2] = Run80000RandomCellsTest();
            timeTable[3] = Run80000RandomCellsTest();
            timeTable[4] = Run80000RandomCellsTest();
            timeTable[5] = Run80000RandomCellsTest();
            timeTable[6] = Run80000RandomCellsTest();
            timeTable[7] = Run80000RandomCellsTest();
            timeTable[8] = Run80000RandomCellsTest();
            timeTable[9] = Run80000RandomCellsTest();

            return timeTable;
        }
        
        private static string GetRandomDate(Random gen)
        {
            DateTime start = new DateTime(1995, 1, 1);
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
        
        private static TimeSpan Run80000RandomCellsTest()
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

            stopwatch.Stop();

            return stopwatch.Elapsed;
        }
    }
}