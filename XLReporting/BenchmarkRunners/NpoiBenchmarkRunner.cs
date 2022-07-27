using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using XLReporting.Configuration;

namespace XLReporting.BenchmarkRunners
{
    internal class NpoiBenchmarkRunner : BenchmarkRunner<ISheet>
    {
        public NpoiBenchmarkRunner(IAppConfig appConfig) : base(appConfig)
        {
        }

        protected override string BenchmarkRunnerName => typeof(NpoiBenchmarkRunner).Name.Replace("BenchmarkRunner", "") ?? "NPOI";
        protected override void PerformBenchmarkWork(Action<ISheet> benchmarkWork, string fileName, bool savingResultingFile)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            
            benchmarkWork(sheet);
            
            if (savingResultingFile)
            {
                workbook.Write(File.Create(fileName));
            }
        }
        protected override void LoadingBigFileBenchmark(bool savingResultingFile)
        {
            var workbook = new XSSFWorkbook(_largeFileName);
            if (savingResultingFile)
            {
                workbook.Write(File.Create(LoadingLargeFileName));
            }
        }
        protected override void CreateRandomCells(ISheet worksheet)
        {
            var rand = new Random();
            for (int i = 1; i <= RandomCellsRowNumber; i++)
            {
                worksheet.CreateRow(i).CreateCell(0).SetCellValue($"=\"{Guid.NewGuid()}\"");
                worksheet.CreateRow(i).CreateCell(1).SetCellValue($"=\"{Guid.NewGuid()}\"");
                worksheet.CreateRow(i).CreateCell(2).SetCellValue(Guid.NewGuid().ToString());
                worksheet.CreateRow(i).CreateCell(3).SetCellValue(rand.Next(32));
                worksheet.CreateRow(i).CreateCell(4).SetCellValue($"=\"{Guid.NewGuid()}\"");
                worksheet.CreateRow(i).CreateCell(5).SetCellValue($"=\"{Guid.NewGuid()}\"");
                worksheet.CreateRow(i).CreateCell(6).SetCellValue(Guid.NewGuid().ToString());
                worksheet.CreateRow(i).CreateCell(7).SetCellValue(rand.Next(13));
                worksheet.CreateRow(i).CreateCell(8).SetCellValue(GetRandomDate(rand));
                worksheet.CreateRow(i).CreateCell(9).SetCellValue(GetRandomDate(rand));
                worksheet.CreateRow(i).CreateCell(10).SetCellValue(Guid.NewGuid().ToString());
                worksheet.CreateRow(i).CreateCell(11).SetCellValue($"=\"{Guid.NewGuid()}\"");
                worksheet.CreateRow(i).CreateCell(12).SetCellValue(Guid.NewGuid().ToString());
                worksheet.CreateRow(i).CreateCell(13).SetCellValue(Guid.NewGuid().ToString());
                worksheet.CreateRow(i).CreateCell(14).SetCellValue((double)GetRandomDecimal(rand));
                worksheet.CreateRow(i).CreateCell(15).SetCellValue((double)GetRandomDecimal(rand));
            }
        }
        protected override void CreateDateCells(ISheet worksheet)
        {
            for (int i = 1; i < DateCellsNumber; i++)
            {
                worksheet.CreateRow(i).CreateCell(1).SetCellValue(DateTime.Now);
            }
        }
        protected override void MakeStyleChanges(ISheet worksheet)
        {
            ICellStyle? style = null;
            for (int i = 1; i <= StyleChangeRowNumber; i++)
            {
                var row = worksheet.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.SetCellValue(_cellValue);

                    if (j == 0)
                    {
                        style = cell.CellStyle;

                        var font = style.GetFont(worksheet.Workbook);
                        font.FontHeight = 22;
                        style.SetFont(font);
                        style.VerticalAlignment = VerticalAlignment.Top;
                        style.Alignment = HorizontalAlignment.Right;
                    }

                    cell.CellStyle = style;
                }
            }
        }
    }
}
