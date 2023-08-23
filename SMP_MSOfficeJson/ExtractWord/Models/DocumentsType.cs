using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractWord.Models
{
    public class DocumentsType
    {
        public List<ContentControlType> contentcontrols = new List<ContentControlType>();// Danh sách tất cả content control( ko nằm trong bảng) trong word
        public List<TableType> tables = new List<TableType>();// Danh sách tất cả các bảng trong word
    }
}
