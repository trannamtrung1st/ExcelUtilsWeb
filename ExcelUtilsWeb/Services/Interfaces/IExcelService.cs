using ExcelUtilsWeb.Models;

namespace ExcelUtilsWeb.Services.Interfaces
{
    public interface IExcelService
    {
        Task<Stream> MergeExcelSheets(Stream excelStream, MergeExcelModel model);
    }
}
