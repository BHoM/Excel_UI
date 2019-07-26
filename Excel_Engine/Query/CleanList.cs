using System;
using System.Collections;
using System.Collections.Generic;
using BH.oM.Excel;
using ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        public static List<T> CleanList<T>(this List<T> list)
        {
            return list.FindAll(item => item != null);
        }
    }
}