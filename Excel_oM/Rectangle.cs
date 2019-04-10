using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.oM.Excel
{
    public class Rectangle : BHoMObject
    {
        public int RowFirst { get; set; }
        public int ColumnFirst { get; set; }
        public int RowLast { get; set; }
        public int ColumnLast { get; set; }
    }
}
