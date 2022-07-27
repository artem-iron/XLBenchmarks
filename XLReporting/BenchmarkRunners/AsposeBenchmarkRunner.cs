
using Aspose.Cells;
using XLReporting.Configuration;

namespace XLReporting.BenchmarkRunners
{
    internal class AsposeBenchmarkRunner : BenchmarkRunner<Cells>
    {
        public AsposeBenchmarkRunner(IAppConfig appConfig) : base(appConfig)
        {
        }

        protected override string BenchmarkRunnerName => typeof(AsposeBenchmarkRunner).Name.Replace("BenchmarkRunner", "") ?? "Aspose";
        public override string NameAndVersion => $"{BenchmarkRunnerName} v.{GetAssemblyVersion(typeof(Cell))}";
        protected override void PerformBenchmarkWork(Action<Cells> benchmarkWork, string fileName, bool savingResultingFile)
        {
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            benchmarkWork(cells);

            if (savingResultingFile)
            {
                workbook.Save(fileName);
            }
        }
        protected override void LoadingBigFileBenchmark(bool savingResultingFile)
        {
            var workbook = new Workbook(_largeFileName);
            if (savingResultingFile)
            {
                workbook.Save(LoadingLargeFileName);
            }
        }
        protected override void CreateRandomCells(Cells cells)
        {
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
        }
        protected override void CreateDateCells(Cells cells)
        {
            var style = new CellsFactory().CreateStyle();
            style.Number = 15;

            for (int i = 1; i < DateCellsNumber; i++)
            {
                var cell = cells["A" + i];
                cell.PutValue(DateTime.Now);
                cell.SetStyle(style);
            }
        }
        protected override void MakeStyleChanges(Cells cells)
        {
            var style = new CellsFactory().CreateStyle();
            style.Font.Size = 22;
            style.VerticalAlignment = TextAlignmentType.Top;
            style.HorizontalAlignment = TextAlignmentType.Right;
            
            cells.InsertRows(1, StyleChangeRowNumber);

            var range = cells.CreateRange($"A1:O{StyleChangeRowNumber}");
            range.Value = _cellValue;

            range.SetStyle(style);
        }
    }
}
