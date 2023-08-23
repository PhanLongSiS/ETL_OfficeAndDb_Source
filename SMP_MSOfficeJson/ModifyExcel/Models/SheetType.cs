using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyExcel.Models
{
    /// <summary>
    ///     Cấu trúc thông tin của một sheet
    /// </summary>
    [JsonObject]
    class SheetType
    {
        /// <summary> Tên của sheet. Vi dụ "STEP 4". </summary>
        public string name;

        /// <summary> Hiện thị hay ẩn sheet. mặc định là hiển thị </summary>

        /// <summary> Danh sách các cell và nội dung </summary>
        /// <see cref="Newtonsoft.Json.JsonSerializationException"> WorkbookData.SheetType[].CellType[]  thì SheetType không cần hàm khởi tạo, nhưng CellType lại băt buộc phải có</see>
        public List<CellType> cells;

        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        public string errMessage = null;

        /// <summary>
        ///         Khai báo thông tin cần ghi vào sheet
        /// </summary>
        /// <param name="cells"> Danh sách nội dung của các cell </param>
        /// <param name="name"> Tên của sheet. Vi dụ "STEP 4"</param>
        /// <param name="visible"> Hiện thị hay ẩn sheet. mặc định là hiển thị</param>
        public SheetType(List<CellType> cells, string name = null, Boolean visible = true)
        {
            this.name = name;
            this.cells = cells;
            errMessage = null;
        }
    }
}
