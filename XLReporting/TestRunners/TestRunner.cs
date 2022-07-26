using System.Diagnostics;
using System.Globalization;
using XLReporting.Configuration;

namespace XLReporting.TestRunners
{
    internal abstract class TestRunner<T>
    {
        public int DateCellsNumber = 80000;
        public int RandomCellsRowNumber = 20000;
        public int StyleChangeRowNumber = 300;

        protected static string _resultsFolderName = "Results";
        protected static string _randomCellsFileNameTemplate = "{0}\\{1}_RandomCells.xlsx";
        protected static string _dateCellsFileNameTemplate = "{0}\\{1}_DateCells.xlsx";
        protected static string _styleChangeFileNameTemplate = "{0}\\{1}_StyleChange.xlsx";
        protected static string _loadingLargeFileFileNameTemplate = "{0}\\{1}_LoadingBigFile.xlsx";
        protected static string _cellValue = "Cell";
        protected static string _largeFileName = "LoadingTestFiles\\LoadingTest.xlsx";

        protected string RandomCellsFileName => string.Format(CultureInfo.InvariantCulture, _randomCellsFileNameTemplate, _resultsFolderName, TestRunnerName);
        protected string DateCellsFileName => string.Format(CultureInfo.InvariantCulture, _dateCellsFileNameTemplate, _resultsFolderName, TestRunnerName);
        protected string StyleChangeFileName => string.Format(CultureInfo.InvariantCulture, _styleChangeFileNameTemplate, _resultsFolderName, TestRunnerName);
        protected string LoadingBigFileName => string.Format(CultureInfo.InvariantCulture, _loadingLargeFileFileNameTemplate, _resultsFolderName, TestRunnerName);


        protected readonly IAppConfig _appConfig;

        public TestRunner(IAppConfig appConfig)
        {
            _appConfig = appConfig;

            DateCellsNumber = _appConfig.DateCellsNumber;
            RandomCellsRowNumber = _appConfig.RandomCellsRowNumber;
            StyleChangeRowNumber = _appConfig.StyleChangeRowNumber;

            _resultsFolderName = _appConfig.ResultsFolderName;
            _randomCellsFileNameTemplate = _appConfig.RandomCellsFileNameTemplate;
            _dateCellsFileNameTemplate = _appConfig.DateCellsFileNameTemplate;
            _styleChangeFileNameTemplate = _appConfig.StyleChangeFileNameTemplate;
            _loadingLargeFileFileNameTemplate = _appConfig.LoadingLargeFileFileNameTemplate;
            _cellValue = _appConfig.CellValue;
            _largeFileName = _appConfig.LargeFileName;
        }

        public TimeSpan[] RunTests()
        {
            CreateResultsFolder();

            var timeTable = new TimeSpan[10];

            timeTable[0] = RunRandomCellsTest(false);
            timeTable[1] = RunRandomCellsTest(true);
            timeTable[2] = RunDateCellsTest(false);
            timeTable[3] = RunDateCellsTest(true);
            timeTable[4] = RunStyleChangesTest(false);
            timeTable[5] = RunStyleChangesTest(true);
            timeTable[6] = RunLoadingBigFileTest(false);
            timeTable[7] = RunLoadingBigFileTest(true);
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
        private TimeSpan RunLoadingBigFileTest(bool savingResultingFile)
        {
            return RunTest(LoadingBigFileTest, savingResultingFile);
        }
        private void RandomCellsTest(bool savingResultingFile)
        {
            DoTestWork(CreateRandomCells, RandomCellsFileName, savingResultingFile);
        }
        private void DateCellsTest(bool savingResultingFile)
        {
            DoTestWork(CreateDateCells, DateCellsFileName, savingResultingFile);
        }
        private void StyleChangesTest(bool savingResultingFile)
        {
            DoTestWork(MakeStyleChanges, StyleChangeFileName, savingResultingFile);
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
        protected abstract void DoTestWork(Action<T> testWork, string fileName, bool savingResultingFile);
        protected abstract void LoadingBigFileTest(bool savingResultingFile);
        protected abstract void CreateRandomCells(T worksheet);
        protected abstract void CreateDateCells(T worksheet);
        protected abstract void MakeStyleChanges(T worksheet);

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

        private static void CreateResultsFolder()
        {
            if (!Directory.Exists(_resultsFolderName))
            {
                Directory.CreateDirectory(_resultsFolderName);
            }
        }
    }
}
