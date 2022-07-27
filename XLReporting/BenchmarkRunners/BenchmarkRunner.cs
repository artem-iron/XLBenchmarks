using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using XLReporting.Configuration;

namespace XLReporting.BenchmarkRunners
{
    internal abstract class BenchmarkRunner<T>
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

        protected string LoadingLargeFileName => string.Format(CultureInfo.InvariantCulture, _loadingLargeFileFileNameTemplate, _resultsFolderName, BenchmarkRunnerName);

        protected readonly IAppConfig _appConfig;

        public BenchmarkRunner(IAppConfig appConfig)
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

        public TimeSpan[] RunBenchmarks()
        {
            CreateResultsFolder();

            var timeTable = new TimeSpan[10];

            timeTable[0] = RunRandomCellsBenchmark(false);
            timeTable[1] = RunRandomCellsBenchmark(true);
            timeTable[2] = RunDateCellsBenchmark(false);
            timeTable[3] = RunDateCellsBenchmark(true);
            timeTable[4] = RunStyleChangesBenchmark(false);
            timeTable[5] = RunStyleChangesBenchmark(true);
            timeTable[6] = RunLoadingBigFileBenchmark(false);
            timeTable[7] = RunLoadingBigFileBenchmark(true);
            timeTable[8] = GetTimeSpan();
            timeTable[9] = GetTimeSpan();

            return timeTable;
        }

        private TimeSpan RunRandomCellsBenchmark(bool savingResultingFile)
        {
            return RunBenchmark(RandomCellsBenchmark, savingResultingFile);
        }
        private TimeSpan RunDateCellsBenchmark(bool savingResultingFile)
        {
            return RunBenchmark(DateCellsBenchmark, savingResultingFile);
        }
        private TimeSpan RunStyleChangesBenchmark(bool savingResultingFile)
        {
            return RunBenchmark(StyleChangesBenchmark, savingResultingFile);
        }
        private TimeSpan RunLoadingBigFileBenchmark(bool savingResultingFile)
        {
            return RunBenchmark(LoadingBigFileBenchmark, savingResultingFile);
        }
        private static TimeSpan RunBenchmark(Action<bool> benchmarkName, bool savingResultingFile)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            benchmarkName(savingResultingFile);

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        private void RandomCellsBenchmark(bool savingResultingFile)
        {
            var randomCellsFileName = string.Format(CultureInfo.InvariantCulture, _randomCellsFileNameTemplate, _resultsFolderName, BenchmarkRunnerName);

            PerformBenchmarkWork(CreateRandomCells, randomCellsFileName, savingResultingFile);
        }
        private void DateCellsBenchmark(bool savingResultingFile)
        {
            var dateCellsFileName = string.Format(CultureInfo.InvariantCulture, _dateCellsFileNameTemplate, _resultsFolderName, BenchmarkRunnerName);

            PerformBenchmarkWork(CreateDateCells, dateCellsFileName, savingResultingFile);
        }
        private void StyleChangesBenchmark(bool savingResultingFile)
        {
            var styleChangeFileName = string.Format(CultureInfo.InvariantCulture, _styleChangeFileNameTemplate, _resultsFolderName, BenchmarkRunnerName);

            PerformBenchmarkWork(MakeStyleChanges, styleChangeFileName, savingResultingFile);
        }

        protected abstract string BenchmarkRunnerName { get; }
        public abstract string NameAndVersion { get; }
        protected abstract void PerformBenchmarkWork(Action<T> benchmarkWork, string fileName, bool savingResultingFile);
        protected abstract void LoadingBigFileBenchmark(bool savingResultingFile);
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

        protected static string GetAssemblyVersion(Type type)
        {
            var assembly = Assembly.GetAssembly(type);
            var assemblyVersion = assembly == null ? null : assembly.GetName().Version;
            var versionString = assemblyVersion == null ? "unknown" : assemblyVersion.ToString();

            return versionString;
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
