using ExcelUtilsWeb.Models;
using ExcelUtilsWeb.Services.Interfaces;
using OfficeOpenXml;

namespace ExcelUtilsWeb.Services
{
    public class ExcelService : IExcelService
    {
        public async Task<Stream> MergeExcelSheets(Stream excelStream, MergeExcelModel model)
        {
            MemoryStream memoryStream = new MemoryStream();

            using (ExcelPackage excel = new ExcelPackage(excelStream))
            {
                var keySet = new HashSet<string>();
                var sheetGroups = new List<List<IGrouping<string, ExcelRange>>>();
                var sheets = model.Sheets.Where(s => !string.IsNullOrEmpty(s)).ToList();
                var columns = model.Columns.Where(c => !string.IsNullOrEmpty(c)).ToList();

                for (int i = 0; i < sheets.Count; i++)
                {
                    string sheetName = sheets[i];
                    string[] columnNames = columns[i].Split('|');

                    ExcelWorksheet sheet = excel.Workbook.Worksheets[sheetName];

                    var startRow = sheet.Dimension.Start.Row;
                    var startCol = sheet.Dimension.Start.Column;
                    var endRow = sheet.Dimension.End.Row;
                    var endCol = sheet.Dimension.End.Column;

                    List<ExcelRange> rows = new List<ExcelRange>();

                    for (var row = startRow; row <= endRow; row++)
                    {
                        rows.Add(sheet.Cells[row, startCol, row, endCol]);
                    }

                    var groups = rows.ToList().GroupBy(r =>
                    {
                        var rowNum = r.Start.Row;
                        var values = string.Join('_', columnNames.Select(c => sheet.Cells[$"{c}{rowNum}"].Value));
                        return values.ToLower();
                    }).Where(g => !string.IsNullOrWhiteSpace(g.Key)).ToList();

                    groups.ForEach(g => keySet.Add(g.Key));

                    sheetGroups.Add(groups);
                }

                var finalSheet = excel.Workbook.Worksheets.Add("_FINAL_");
                var currentCol = 1;
                var currentRow = 1;

                foreach (var key in keySet)
                {
                    if (string.Equals("S22A04983", key, StringComparison.OrdinalIgnoreCase))
                    {

                    }

                    var maxRow = 0;

                    foreach (var sheetGroup in sheetGroups)
                    {
                        var group = sheetGroup.Where(g => g.Key == key).FirstOrDefault();

                        if (group != null)
                        {
                            var totalRows = group.Count();
                            var maxCol = 0;

                            for (var i = 0; i < totalRows; i++)
                            {
                                var row = currentRow + i;

                                var range = group.ElementAt(i);

                                range.Copy(finalSheet.Cells[row, currentCol, row, currentCol + range.Columns]);

                                if (range.Columns > maxCol)
                                {
                                    maxCol = range.Columns;
                                }
                            }

                            currentCol += maxCol;

                            if (currentRow + totalRows > maxRow)
                            {
                                maxRow = currentRow + totalRows;
                            }
                        }
                    }

                    currentRow = maxRow;
                    currentCol = 1;
                }

                await excel.SaveAsAsync(memoryStream);
            }

            memoryStream.Seek(0, SeekOrigin.Begin);

            return memoryStream;
        }
    }
}
