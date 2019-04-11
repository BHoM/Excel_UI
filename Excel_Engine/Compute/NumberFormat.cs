using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel
{
    public static partial class Compute
    {
        public static bool NumberFormat(this oM.Excel.Reference reference, string fmt = null)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    ExcelReference selected = XlCall.Excel(XlCall.xlfSelection) as ExcelReference;
                    XlCall.Excel(XlCall.xlcSelect, reference.ToExcel());
                    XlCall.Excel(XlCall.xlcFormatNumber, fmt);
                    XlCall.Excel(XlCall.xlcSelect, selected);
                }
                catch (XlCallException exception)
                {
                    Reflection.Compute.RecordError(exception.Message);
                }
            });
            return true;
        }
    }
}
