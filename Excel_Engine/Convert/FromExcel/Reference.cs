using BH.oM.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel
{
    public static partial class Convert
    {
        public static Reference ToBHoM(this ExcelDna.Integration.ExcelReference xlRef)
        {
            return Create.Reference(xlRef.RowFirst, xlRef.RowLast, xlRef.ColumnFirst, xlRef.ColumnLast, xlRef.SheetId);
        }
    }
}
