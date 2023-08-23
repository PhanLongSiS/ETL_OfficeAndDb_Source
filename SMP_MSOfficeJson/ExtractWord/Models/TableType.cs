using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractWord.Models
{
    public class TableType
    {
        public string name { get; set; }
        public List<RowValue> rowvalues = new List<RowValue>();
        public string erroMess { get; set; } = null;
    }
}
