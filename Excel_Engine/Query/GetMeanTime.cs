using System;
using System.Collections.Generic;
using BH.oM.Excel;
using ExcelDna.Integration;

namespace BH.Engine.Excel.Profiling

{
    public static partial class Query
    {
        public static double GetMeanTime(string name)
        {
            return Timer.GetMean(name);
        }
    }
}