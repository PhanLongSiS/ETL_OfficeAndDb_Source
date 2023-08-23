using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformExJsToWordJs.ModelWordJson
{
    public class ContentControlType
    {
        /// <summary>
        /// Thông tin từng dữ liệu trong Word 
        /// Ví dụ:Họ và tên:Nguyễn Đức Tiến được lưu trong file json {ContentTitle:Họ và Tên ; ContentText:Nguyễn Đức Tiến}
        /// </summary>
        public string title { get; set; } // Title của Control
        public string value { get; set; }// Text của Control
        public string erroMess { get; set; } = null;
    }
}
