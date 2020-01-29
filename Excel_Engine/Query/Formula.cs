using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        public static string Formula(this oM.Excel.Reference reference)
        {
            return XlCall.Excel(XlCall.xlfGetFormula, reference.ToExcel()).ToString();
        }
    }
}