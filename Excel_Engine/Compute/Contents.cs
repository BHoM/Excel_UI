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
        public static bool Contents(this oM.Excel.Reference reference, string value)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    XlCall.Excel(XlCall.xlcFormula, value, reference.ToExcel());
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
