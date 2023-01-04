/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *
 *
 * The BHoM is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3.0 of the License, or
 * (at your option) any later version.
 *
 * The BHoM is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.
 */

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.UI.Excel.Templates;
using Microsoft.Office.Interop.Excel;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static void Old_Restore()
        {
            foreach (var restored in Old_GetFormulas())
            {
                string json = restored.Item2;
                string callerType = restored.Item1;
                if (AddIn.CallerShells.ContainsKey(callerType))
                {
                    // Register that formula from the json information
                    CallerFormula formula = InstantiateCaller(callerType);
                    if (formula != null)
                    {
                        formula.Caller.Read(json);
                        Register(formula);
                    }
                }
            }
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static List<Tuple<string, string>> Old_GetFormulas()
        {
            // Get the old hidden page for storing the callers
            Worksheet sheet = Sheet("BHoM_ComponetRequests", false);
            if (sheet == null)
                return new List<Tuple<string, string>>();

            // Collect the formulas
            List<Tuple<string, string>> components = new List<Tuple<string, string>>();
            foreach (Range row in sheet.UsedRange)
            {
                string str = "";
                string callerType = "";
                try
                {
                    Range cell = row.Cells[1, 2];
                    callerType = cell.Value.ToString();

                    int col = 3;
                    cell = row.Cells[1, col++];
                    while (cell.Value != null && cell.Value is string && (cell.Value as string).Length > 0)
                    {
                        str += cell.Value;
                        cell = row.Cells[1, col++];
                    }
                }
                catch { }

                if (str.Length > 0)
                    components.Add(new Tuple<string, string>(callerType, str));
            }

            return components;
        }

        /*******************************************/
    }
}




