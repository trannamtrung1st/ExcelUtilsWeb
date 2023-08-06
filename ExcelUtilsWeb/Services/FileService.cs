using ExcelUtilsWeb.Services.Interfaces;
using Microsoft.JSInterop;

namespace ExcelUtilsWeb.Services
{
    public class FileService : IFileService
    {
        private readonly IJSRuntime _jsRuntime;

        public FileService(IJSRuntime jsRuntime)
        {
            _jsRuntime = jsRuntime;
        }

        public async Task DownloadFileFromStream(Stream stream, string fileName)
        {
            using DotNetStreamReference streamRef = new DotNetStreamReference(stream);

            await _jsRuntime.InvokeVoidAsync("downloadFileFromStream", fileName, streamRef);
        }

    }
}
