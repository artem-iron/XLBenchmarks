using IronXL;
using IronXL.Styles;

namespace CurrentIronXL
{
    public class TestRunner : TestRunnerBase.TestRunner<WorkSheet>
    {
        protected override string TestRunnerName => typeof(TestRunner).Namespace ?? "CurrentIronXL";
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
            var workbook = WorkBook.Load(BigFileName);
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
            for (int i = 1; i < DateCellsNumber; i++)
            {
                worksheet["A" + i].Value = DateTime.Now;
            }
        }
        protected override void MakeStyleChanges(WorkSheet worksheet)
        {
            worksheet.InsertRows(1, StyleChangeRowNumber);

            var range = worksheet.GetRange($"A1:O{StyleChangeRowNumber}");
            range.Value = CELL_VALUE;

            var style = range.Style;

            style.Font.Height = 22;
            style.VerticalAlignment = VerticalAlignment.Top;
            style.HorizontalAlignment = HorizontalAlignment.Right;
        }
    }
}