/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2021, the respective contributors. All rights reserved.
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

using BH.oM.Adapters.Excel;
using BH.oM.Reflection.Attributes;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Creates a BHoM CellAddress based on the given programmatic column and row indices, starting from column index 0 equal to A in Excel and row index 0 equal to 1 in Excel.")]
        [Input("columnIndex", "Index of the column to be used, starting from 0 equal to A in Excel.")]
        [Input("rowIndex", "Index of the row to be used, starting from 0 equal to 1 in Excel.")]
        [Output("column", "Letter representing the column under given index.")]
        public static CellAddress CellAddress(int columnIndex, int rowIndex)
        {
            string column = columnIndex.ColumnName();
            if (string.IsNullOrWhiteSpace(column))
                return null;

            if (rowIndex < 0)
            {
                BH.Engine.Reflection.Compute.RecordError("Negative row index is not allowed.");
                return null;
            }

            return new CellAddress { Column = column, Row = rowIndex + 1 };
        }

        /*******************************************/

        [Description("Creates a BHoM CellAddress based on the given string representing cell address in Excel-readable format.")]
        [Input("excelAddress", "String representing cell address in Excel-readable format.")]
        [Output("address", "BHoM CellAddress object created based on the input string.")]
        public static CellAddress CellAddress(string excelAddress)
        {
            if (!excelAddress.IsValidAddress())
                return null;

            string column = Regex.Match(excelAddress, @"[A-Z]+").Value;
            int row = int.Parse(Regex.Match(excelAddress, @"\d+").Value);

            return new CellAddress { Column = column, Row = row };
        }

        /*******************************************/
    }
}
