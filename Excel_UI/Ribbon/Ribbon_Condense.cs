/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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

using BH.UI.Excel.Templates;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BH.UI.Excel.Addin
{
    public partial class Ribbon : ExcelRibbon
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void RunCondense(IRibbonControl control)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddIn.WriteFormula("=BHoM.Condense");
            });
        }

        /*******************************************/

        [ExcelFunction(Name = "BHoM.Condense", Description = "Take a group of cells and store their content as a list in a single cell.", Category = "UI")]
        public static object Condense(object item)
        {
            object result = new object();
            if (item is object[,] array
                && (array.GetLength(0) == 1 || array.GetLength(1) == 1))
            {
                var filteredItems = array
                    .Cast<object>()
                    .Where(x => !(x is ExcelEmpty))
                    .ToArray();
                result = AddIn.FromExcel(filteredItems).ToList();
                return AddIn.ToExcel(result);
            }

            if (item is object[,] matrix
                && matrix.GetLength(0) > 1
                && matrix.GetLength(1) > 1)
            {
                var listResult = new List<object>();

                for (int i = 0; i < matrix.GetLength(0); i++)
                {
                    List<object> row = Enumerable
                        .Range(0, matrix.GetLength(1))
                        .Select(x => matrix[i, x]).ToList();
                    listResult.Add(AddIn.FromExcel(row));
                }

                result = AddIn.FromExcel(listResult);
                return AddIn.ToExcel(result);
            }

            else
            {
                result = AddIn.FromExcel(item);
                return AddIn.ToExcel(result);
            }

        }

        /*******************************************/
    }
}





