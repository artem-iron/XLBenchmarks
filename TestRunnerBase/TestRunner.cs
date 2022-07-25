using System.Globalization;

namespace TestRunnerBase
{
    public abstract class TestRunner
    {
        public static readonly int DateCellsNumber = 80000;
        public static readonly int RandomCellsRowNumber = 20000;
        public static readonly int StyleChangeRowNumber = 300;

        public static readonly string RANDOM_CELLS_FILE_NAME_TEMPLATE = "{0}_RandomCells.xlsx";
        public static readonly string DATE_CELLS_FILE_NAME_TEMPLATE = "{0}_DateCells.xlsx";
        public static readonly string STYLE_CHANGE_FILE_NAME_TEMPLATE = "{0}_StyleChange.xlsx";
        public static readonly string CELL_VALUE = "CellValue";

        public string RandomCellsFileName => string.Format(CultureInfo.InvariantCulture, RANDOM_CELLS_FILE_NAME_TEMPLATE, TestRunnerName);
        public string DateCellsFileName => string.Format(CultureInfo.InvariantCulture, DATE_CELLS_FILE_NAME_TEMPLATE, TestRunnerName);
        public string StyleChangeFileName => string.Format(CultureInfo.InvariantCulture, STYLE_CHANGE_FILE_NAME_TEMPLATE, TestRunnerName);

        public TimeSpan[] RunTests()
        {
            var timeTable = new TimeSpan[10];

            timeTable[0] = RunRandomCellsTest(false);
            timeTable[1] = RunRandomCellsTest(true);
            timeTable[2] = RunDateCellsTest(false);
            timeTable[3] = RunDateCellsTest(true);
            timeTable[4] = RunStyleChangesTest(false);
            timeTable[5] = RunStyleChangesTest(true);
            timeTable[6] = GetTimeSpan();
            timeTable[7] = GetTimeSpan();
            timeTable[8] = GetTimeSpan();
            timeTable[9] = GetTimeSpan();

            return timeTable;
        }

        public abstract string TestRunnerName { get; }
        public abstract TimeSpan RunRandomCellsTest(bool savingResultingFile);
        public abstract TimeSpan RunDateCellsTest(bool savingResultingFile);
        public abstract TimeSpan RunStyleChangesTest(bool savingResultingFile);


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