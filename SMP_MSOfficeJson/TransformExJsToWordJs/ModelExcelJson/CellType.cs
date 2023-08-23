using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformExJsToWordJs.ModelExcelJson
{
    /// <summary>
    ///     Cấu trúc thông tin của một cell
    ///     Vi du (dong2, cot3)
    ///     Chu y con dang dia chi A5, AddressHint
    /// </summary>
    public class CellType
    {
        public string posName { get; set; }
        public string pos { get; set; }
        public string value { get; set; }
    }
}
