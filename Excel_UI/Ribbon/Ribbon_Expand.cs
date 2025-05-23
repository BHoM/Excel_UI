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

using BH.Engine.Base;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BH.UI.Excel.Addin
{
    public partial class Ribbon : ExcelRibbon
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void RunExpand(IRibbonControl control)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                AddIn.WriteFormula("=BHoM.Expand");
            });
        }

        /*******************************************/

        [ExcelFunction(Name = "BHoM.Expand", Description = "Take a list stored in a single cell and expand over multiple cells (one cell per item in the list).", Category = "UI")]
        public static object Expand(object item, bool transpose = false, bool onlyExpandFirstDimension = false)
        {
            item = AddIn.FromExcel(item);
            object[,] result;

            if (item is IEnumerable<object> array)
            {
                List<object> content = array.Cast<object>().ToList();
                if (!onlyExpandFirstDimension && content.All(x => x is IEnumerable<object>))
                {
                    result = AddIn.ToExcel(content.OfType<IEnumerable>().Select(x => x.OfType<object>().ToList()).ToList());
                }  
                else
                {
                    int count = content.Count();
                    result = new object[count, 1];
                    for (int i = 0; i < count; i++)
                        result[i, 0] = AddIn.ToExcel(content[i]);
                }
            }
            else
            {
                result = new object[,] { { item } };
            }

            if (transpose)
            {
                object[,] transposed = new object[result.GetLength(1), result.GetLength(0)];
                for (int i = 0; i < result.GetLength(0); i++)
                {
                    for (int j = 0; j < result.GetLength(1); j++)
                        transposed[j, i] = result[i, j];
                }
                return AddIn.ToExcel(transposed);
            }

            return AddIn.ToExcel(result);
        }

        /*******************************************/
    }
}





