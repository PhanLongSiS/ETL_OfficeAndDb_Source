using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyExcel.Models
{
    /// <summary>
    ///     Cấu trúc thông tin của một cell
    ///     Vi du (dong2, cot3)
    ///     Chu y con dang dia chi A5, AddressHint
    /// </summary>
    class CellType : CellPosition
    {

        private string _pos;

        /// <summary> Vị trí của cell. Vi dụ A1, C4 </summary>
        [JsonProperty]
        public string pos
        {
            set
            {
                CellPosition.StringAddressToNumber(value, ref this.ColumnIndex, ref this.RowIndex);
                _pos = value;
            }
            get
            {
                return _pos;
            }
        }


        /// <summary> Dữ liệu của cell </summary>
        [JsonProperty]
        public object value;
        public string posname;

        /// <summary>
        ///      Hàm khởi tạo 
        /// </summary>
        /// <see cref="Newtonsoft.Json.JsonSerializationException"> WorkbookData.SheetType.CellType  thì SheetType không cần hàm khởi tạo, nhưng CellType lại băt buộc phải có</see>
        public CellType()
        {
            // Must have, eventhough the body is empty.
        }

        public CellType(string pos, object value)
        {
            this.pos = pos;
            this.value = value;
        }

        public CellType(KeyValuePair<string, object> item)
        {
            pos = item.Key;
            value = item.Value;
        }
    }
}
