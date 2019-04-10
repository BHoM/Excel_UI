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
        public static double ToExcel(this DateTime dateTime)
        {
            return dateTime.ToOADate();
        }
    }
}
