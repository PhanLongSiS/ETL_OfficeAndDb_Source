using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyExcel.Models
{
    /// <summary>
    ///     Cấu trúc dữ liệu cần đưa vào một file excel
    /// </summary>
    class WorkbookType
    {
        /// <summary> Cấu hình workbook </summary>
        public WorkbookConfig config;
        /// <summary> Nội dung của các sheet cần đẩy vào file excel </summary>
        public List<SheetType> sheets;

        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        public string errMessage = "";

        public WorkbookType()
        {
            config = new WorkbookConfig();
            sheets = new List<SheetType>();
        }

    }
}
