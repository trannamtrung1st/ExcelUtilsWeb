using Microsoft.AspNetCore.Components.Forms;

namespace ExcelUtilsWeb.Models
{
    public class MergeExcelModel
    {
        public MergeExcelModel()
        {
            Sheets = new List<string>();
            Columns = new List<string>();
        }

        public IBrowserFile File { get; set; }
        public List<string> Sheets { get; set; }
        public List<string> Columns { get; set; }
    }
}
