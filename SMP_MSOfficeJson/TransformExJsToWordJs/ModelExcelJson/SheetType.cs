using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformExJsToWordJs.ModelExcelJson
{
    public class SheetType
    {

        /// <summary> Tên của sheet. Vi dụ "STEP 4". </summary>
        public string name { get; set; }

        ///// <summary> Hiện thị hay ẩn sheet. mặc định là hiển thị </summary>
        //public string visible { get; set; }

        /// <summary> Danh sách các cell và nội dung </summary>
        /// <see cref="Newtonsoft.Json.JsonSerializationException"> WorkbookData.SheetType[].CellType[]  thì SheetType không cần hàm khởi tạo, nhưng CellType lại băt buộc phải có</see>
        public List<CellType> cells { get; set; } = new List<CellType>();
        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        public string errMessage = null;
    }
}
