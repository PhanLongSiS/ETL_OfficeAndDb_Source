using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyWord.Models
{
    /// <summary>
    ///     Cấu trúc dữ liệu cần đưa vào một file excel
    /// </summary>
    class WorkspaceType
    {
        /// <summary> Cấu hình workbook </summary>
        public WorkspaceConfig config;
        /// <summary> Nội dung của các documents cần đẩy vào file word </summary>
        public DocumentType documents;
        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        public string errMessage = "";
    }
}
