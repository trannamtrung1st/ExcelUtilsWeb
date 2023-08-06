namespace ExcelUtilsWeb.Services.Interfaces
{
    public interface IFileService
    {
        Task DownloadFileFromStream(Stream stream, string fileName);
    }
}