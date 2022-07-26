﻿using IronXLOld;
using IronXLOld.Styles;
using XLReporting.Configuration;

namespace XLReporting.TestRunners
{
    internal class PreviousIronXLTestRunner : TestRunner<WorkSheet>
    {
        public PreviousIronXLTestRunner(IAppConfig appConfig) : base(appConfig)
        {
        }

        protected override string TestRunnerName => typeof(PreviousIronXLTestRunner).Name.Replace("TestRunner", "") ?? "PreviousIronXL";
        protected override void DoTestWork(Action<WorkSheet> testWork, string fileName, bool savingResultingFile)
        {
            var workbook = new WorkBook();
            var cells = workbook.DefaultWorkSheet;

            testWork(cells);

            if (savingResultingFile)
            {
                workbook.SaveAs(fileName);
            }
        }
        protected override void LoadingBigFileTest(bool savingResultingFile)
        {
            var workbook = WorkBook.Load(_largeFileName);
            if (savingResultingFile)
            {
                workbook.SaveAs(LoadingBigFileName);
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
            int dateCellsNumber = 60000;
            for (int i = 1; i < dateCellsNumber; i++)
            {
                worksheet["A" + i].Value = DateTime.Now;
            }
        }
        protected override void MakeStyleChanges(WorkSheet worksheet)
        {
            int styleChangeRowNumber = 50;
            worksheet.InsertRows(1, styleChangeRowNumber);

            var range = worksheet.GetRange($"A1:A{styleChangeRowNumber}");
            range.Value = _cellValue;

            var style = range.Style;

            style.Font.Height = 22;
            style.VerticalAlignment = VerticalAlignment.Top;
            style.HorizontalAlignment = HorizontalAlignment.Right;
        }
    }
}