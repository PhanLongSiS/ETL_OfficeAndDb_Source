using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModifyWord.Models
{
    /// <summary>
    ///         Cầu hình cho thông tin cần đẩy vào wookbook
    /// </summary>
    /// <example>
    ///     - Attribute [JsonObject] chỉ có tác dụng với Serialize, không có tác dụng với Deserialize
    /// </example>
    [JsonObject]
    class WorkspaceConfig
    {
        /// <summary> Có cho phép nhìn thấy tiến trình excel đang chạy không? </summary>
        public bool visible { get; set; }

        /// <summary> Có cho phép in tất cả các sheet (phụ thuộc vào cấu hình của từng sheet nữa) ra máy in mặc định không? </summary>
        /// <seealso cref="printername"/>
        public bool printnow { get; set; }

        /// <summary> Tên của máy in muốn xuất. Chọn máy in đầu tiên phù hợp nếu tên chung chung như là pdf. </summary>
        /// <seealso cref="printnow"/>
        public string printername { get; set; }

        /// <summary> Tắt ngay excel sau khi quá trình thực hiện kết thúc? </summary>
        public bool terminate { get; set; }

        /// <summary> Tên file muốn lưu lại sau khi đã đổ số liệu. Bỏ qua nếu không muốn ghi ra file </summary>
        public string saveas { get; set; }

        public WorkspaceConfig()
        {
            visible = false;
            printnow = false;
            terminate = true;
            saveas = string.Empty;
            printername = string.Empty;
        }

    }
}
