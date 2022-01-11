/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2022, the respective contributors. All rights reserved.
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
using BH.oM.Base.Attributes;
using System.ComponentModel;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Creates a BHoM CellRange based on start and end cells in an Excel-readable string format.")]
        [Input("from", "Top-left corner of the range in a Excel-readable string format.")]
        [Input("to", "Bottom-right corner of the range in a Excel-readable string format.")]
        [Output("range", "BHoM CellRange created based on the input strings.")]
        public static CellRange CellRange(string from, string to)
        {
            return Create.CellRange($"{from}:{to}");
        }

        /*******************************************/

        [Description("Creates a BHoM CellRange based on the given string representing cell range in Excel-readable format.")]
        [Input("excelRange", "String representing cell range in Excel-readable format.")]
        [Output("range", "BHoM CellRange object created based on the input string.")]
        public static CellRange CellRange(string excelRange)
        {
            if (!excelRange.IsValidRange())
                return null;

            string[] split = excelRange.Split(new char[] { ':' });
            string from = split[0];
            string to = split[1];

            return new CellRange { From = Create.CellAddress(from), To = Create.CellAddress(to) };
        }

        /*******************************************/
    }
}

