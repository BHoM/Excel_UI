using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Convert
    {
        public static ExcelReference ToExcel(this oM.Excel.Reference omRef)
        {
            var rects = omRef.Rectangles.Select((rect) =>
                new int[] { rect.RowFirst, rect.RowLast, rect.ColumnFirst, rect.ColumnLast }).ToArray();
            return new ExcelReference(rects, omRef.Sheet);
        }
    }
}
