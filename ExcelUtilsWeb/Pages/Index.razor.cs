using AntDesign;
using ExcelUtilsWeb.Models;
using ExcelUtilsWeb.Services.Interfaces;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;

namespace ExcelUtilsWeb.Pages
{
    public partial class Index
    {
        [Inject]
        IMessageService Message { get; set; }

        [Inject]
        IExcelService ExcelService { get; set; }

        [Inject]
        IFileService FileService { get; set; }

        MergeExcelModel model { get; set; }
        bool loading { get; set; }

        public Index()
        {
            model = new MergeExcelModel();
            AddNewSheet();
        }

        void OnChange(string value, List<string> list, int index)
        {
            list[index] = value;
            var lastSheet = model.Sheets.Last();
            var lastColumn = model.Columns.Last();

            if (!string.IsNullOrEmpty(lastSheet) && !string.IsNullOrEmpty(lastColumn))
            {
                AddNewSheet();
            }
        }

        void AddNewSheet()
        {
            model.Sheets.Add("");
            model.Columns.Add("");
        }

        void OnFileChanged(InputFileChangeEventArgs args)
        {
            model.File = args.File;
        }

        async Task OnFinish(EditContext context)
        {
            try
            {
                loading = true;

                MergeExcelModel model = context.Model as MergeExcelModel;

                using Stream uploadedStream = model.File.OpenReadStream(maxAllowedSize: long.MaxValue);
                using MemoryStream memStream = new MemoryStream();
                await uploadedStream.CopyToAsync(memStream);
                memStream.Seek(0, SeekOrigin.Begin);

                Stream mergedFileStream = await ExcelService.MergeExcelSheets(memStream, model);

                await FileService.DownloadFileFromStream(mergedFileStream, model.File.Name);

                _ = Message.Success("Kết nối thành công!");

                loading = false;
            }
            catch (Exception ex)
            {
                loading = false;

                Console.Error.WriteLine(ex);

                _ = Message.Error("Có lỗi xảy ra!");
            }
        }

        void OnFinishFailed(EditContext context)
        {
            Console.WriteLine("Không hợp lệ");
        }
    }
}
