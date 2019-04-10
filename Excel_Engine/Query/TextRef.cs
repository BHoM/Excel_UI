using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        public static string RefText(this oM.Excel.Reference reference)
        {
            return XlCall.Excel(XlCall.xlfReftext, reference.ToExcel(), true).ToString();
        }
    }
}