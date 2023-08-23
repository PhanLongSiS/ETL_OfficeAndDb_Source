using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyWord.Models
{
    /// <summary>
    ///     Cấu trúc thông tin của một content type 
    /// </summary>
    class ContentControlType
    {

        private string _title;

        /// <summary> Tên của content type. Vi dụ mssv </summary>
        [JsonProperty]
        public string title
        {
            set
            {
                _title = value;
            }
            get
            {
                return _title;
            }
        }


        /// <summary> Dữ liệu của cell </summary>
        [JsonProperty]
        public object value;

        /// <summary>
        ///      Hàm khởi tạo 
        /// </summary>
        /// <see cref="Newtonsoft.Json.JsonSerializationException"> WorkbookData.SheetType.CellType  thì SheetType không cần hàm khởi tạo, nhưng CellType lại băt buộc phải có</see>
        public ContentControlType()
        {
            // Must have, eventhough the body is empty.
        }

        public ContentControlType(string pos, object value)
        {
            this.title = pos;
            this.value = value;
        }

        public ContentControlType(KeyValuePair<string, object> item)
        {
            title = item.Key;
            value = item.Value;
        }
    }
}
