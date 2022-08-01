﻿using IronXLOld;
using IronXLOld.Styles;
using XLBenchmarks.Configuration;

namespace XLBenchmarks.BenchmarkRunners
{
    internal class PreviousIronXLBenchmarkRunner : BenchmarkRunner<WorkSheet>
    {
        public PreviousIronXLBenchmarkRunner(IAppConfig appConfig) : base(appConfig)
        {
        }

        protected override string BenchmarkRunnerName => typeof(PreviousIronXLBenchmarkRunner).Name.Replace("BenchmarkRunner", "") ?? "PreviousIronXL";
        public override string NameAndVersion => $"{BenchmarkRunnerName} v.{GetAssemblyVersion(typeof(Cell))}";
        protected override void PerformBenchmarkWork(Action<WorkSheet> benchmarkWork, string fileName, bool savingResultingFile)
        {
            var workbook = new WorkBook();
            var cells = workbook.DefaultWorkSheet;

            benchmarkWork(cells);

            if (savingResultingFile)
            {
                workbook.SaveAs(fileName);
            }
        }
        protected override void LoadingBigFileBenchmark(bool savingResultingFile)
        {
            var workbook = WorkBook.Load(_largeFileName);
            if (savingResultingFile)
            {
                workbook.SaveAs(LoadingLargeFileName);
            }
        }
        protected override void CreateRandomCells(WorkSheet worksheet)
        {
            var rand = new Random();
            for (int i = 1; i <= RandomCellsRowNumber; i++)
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
        }
        protected override void CreateDateCells(WorkSheet worksheet)
        {
            for (int i = 1; i < DateCellsNumber; i++)
            {
                worksheet["A" + i].Value = DateTime.Now;
            }
        }
        protected override void MakeStyleChanges(WorkSheet worksheet)
        {
            worksheet.InsertRows(1, StyleChangeRowNumber);

            var range = worksheet.GetRange($"A1:A{StyleChangeRowNumber}");
            range.Value = _cellValue;

            var style = range.Style;

            style.Font.Height = 22;
            style.VerticalAlignment = VerticalAlignment.Top;
            style.HorizontalAlignment = HorizontalAlignment.Right;
        }

        protected override void GenerateFormulas(WorkSheet worksheet)
        {
            var rnd = new Random();

            for (int i = 1; i <= GenerateFormulasRowNumber; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    string cellA = $"{_letters[rnd.Next(1, 10)]}{rnd.Next(GenerateFormulasRowNumber + 1, GenerateFormulasRowNumber * 2)}";
                    string cellB = $"{_letters[rnd.Next(1, 10)]}{rnd.Next(GenerateFormulasRowNumber + 1, GenerateFormulasRowNumber * 2)}";
                    worksheet[$"{_letters[j]}{i}"].Formula = $"={cellA}/{cellB}";
                }
            }

            for (int i = GenerateFormulasRowNumber + 1; i <= GenerateFormulasRowNumber * 2; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    worksheet[$"{_letters[j]}{i}"].Value = GetRandomRandInt(rnd);
                }
            }
        }
    }
}
