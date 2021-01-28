using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    [Serializable]
    public class ExcelRow
    {
        public Guid Guid { get; set; }
        public int RowNumber { get; set; }
        public List<int> Data { get; set; }
        public bool y { get; set; }
        public bool y2 { get; set; }
    }
}
