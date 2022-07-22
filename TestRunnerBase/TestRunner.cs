using System.Globalization;

namespace TestRunnerBase
{
    public abstract class TestRunner
    {
        public TimeSpan[] RunTests()
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

        public abstract TimeSpan Run320000RandomCellsTest();
        public abstract TimeSpan Run160000DateCellsTest();
        public abstract TimeSpan RunStyleChangesTest();


        public static TimeSpan GetTimeSpan()
        {
            return TimeSpan.FromSeconds(10);
        }

        public static string GetRandomDate(Random gen)
        {
            DateTime start = new(1995, 1, 1);
            int range = (DateTime.Today - start).Days;
            return start.AddDays(gen.Next(range)).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        public static decimal GetRandomDecimal(Random rng)
        {
            byte scale = (byte)rng.Next(29);
            bool sign = rng.Next(2) == 1;
            return new decimal(GetRandomRandInt(rng),
                GetRandomRandInt(rng),
                GetRandomRandInt(rng),
                sign,
                scale);
        }

        public static int GetRandomRandInt(Random rng)
        {
            int firstBits = rng.Next(0, 1 << 4) << 28;
            int lastBits = rng.Next(0, 1 << 28);
            return firstBits | lastBits;
        }
    }
}