using System.Diagnostics;
using System.Globalization;

namespace TestRunnerBase
{
    public abstract class TestRunner
    {
        protected static readonly int DateCellsNumber = 80000;
        protected static readonly int RandomCellsRowNumber = 20000;
        protected static readonly int StyleChangeRowNumber = 300;

        protected static readonly string RANDOM_CELLS_FILE_NAME_TEMPLATE = "{0}_RandomCells.xlsx";
        protected static readonly string DATE_CELLS_FILE_NAME_TEMPLATE = "{0}_DateCells.xlsx";
        protected static readonly string STYLE_CHANGE_FILE_NAME_TEMPLATE = "{0}_StyleChange.xlsx";
        protected static readonly string CELL_VALUE = "CellValue";

        protected string RandomCellsFileName => string.Format(CultureInfo.InvariantCulture, RANDOM_CELLS_FILE_NAME_TEMPLATE, TestRunnerName);
        protected string DateCellsFileName => string.Format(CultureInfo.InvariantCulture, DATE_CELLS_FILE_NAME_TEMPLATE, TestRunnerName);
        protected string StyleChangeFileName => string.Format(CultureInfo.InvariantCulture, STYLE_CHANGE_FILE_NAME_TEMPLATE, TestRunnerName);

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

        private TimeSpan RunRandomCellsTest(bool savingResultingFile)
        {
            return RunTest(RandomCellsTest, savingResultingFile);
        }
        private TimeSpan RunDateCellsTest(bool savingResultingFile)
        {
            return RunTest(DateCellsTest, savingResultingFile);
        }
        private TimeSpan RunStyleChangesTest(bool savingResultingFile)
        {
            return RunTest(StyleChangesTest, savingResultingFile);
        }
        private static TimeSpan RunTest(Action<bool> testName, bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            testName(savingResultingFile);
            
            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
        
        protected abstract string TestRunnerName { get; }
        protected abstract void RandomCellsTest(bool savingResultingFile);
        protected abstract void DateCellsTest(bool savingResultingFile);
        protected abstract void StyleChangesTest(bool savingResultingFile);


        protected static TimeSpan GetTimeSpan()
        {
            return TimeSpan.FromSeconds(10);
        }

        protected static string GetRandomDate(Random gen)
        {
            DateTime start = new(1995, 1, 1);
            int range = (DateTime.Today - start).Days;
            return start.AddDays(gen.Next(range)).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        protected static decimal GetRandomDecimal(Random rng)
        {
            byte scale = (byte)rng.Next(29);
            bool sign = rng.Next(2) == 1;
            return new decimal(GetRandomRandInt(rng),
                GetRandomRandInt(rng),
                GetRandomRandInt(rng),
                sign,
                scale);
        }

        protected static int GetRandomRandInt(Random rng)
        {
            int firstBits = rng.Next(0, 1 << 4) << 28;
            int lastBits = rng.Next(0, 1 << 28);
            return firstBits | lastBits;
        }
    }
}