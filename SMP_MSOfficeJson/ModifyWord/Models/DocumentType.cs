using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyWord.Models
{
    /// <summary>
    ///     Cấu trúc thông tin của một document
    /// </summary>
    [JsonObject]
    class DocumentType
    {
        /// <summary> Danh sách các cell và nội dung </summary>
        /// <see cref="Newtonsoft.Json.JsonSerializationException"> WorkbookData.SheetType[].CellType[]  thì SheetType không cần hàm khởi tạo, nhưng CellType lại băt buộc phải có</see>
        public List<ContentControlType> contentcontrols;

        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        public string errMessage = null;

        /// <summary>
        ///         Khai báo thông tin cần ghi vào content control
        /// </summary>
        /// <param name="controls"> Danh sách nội dung của các content controls </param>
        /// <param name="name"> Tên của sheet. Vi dụ "STEP 4"</param>
        /// <param name="visible"> Hiện thị hay ẩn sheet. mặc định là hiển thị</param>
        public DocumentType(List<ContentControlType> controls)
        {
            this.contentcontrols = controls;
            errMessage = null;
        }
    }
}
