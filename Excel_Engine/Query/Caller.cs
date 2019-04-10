using System;
using System.Collections.Generic;
using BH.oM.Excel;
using ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        public static oM.Excel.Reference Caller()
        {
            ExcelReference xlref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            return xlref.ToBHoM();
        }
    }
}