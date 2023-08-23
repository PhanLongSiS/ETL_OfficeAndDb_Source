using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformExJsToWordJs.ModelExcelJson
{
    public class WorkbookType
    {
        public WorkbookConfig config { get; set; }
        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        /// <summary> Nội dung của các sheet cần đẩy vào file excel </summary>
        public List<SheetType> sheets { get; set; } = new List<SheetType>();
        /// <summary> Cấu hình workbook </summary>
        public string errMessage = "";
        public WorkbookType()
        {
            config = new WorkbookConfig();
            sheets = new List<SheetType>();
        }
    }
}
