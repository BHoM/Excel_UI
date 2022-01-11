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

        [Description("Creates an ExcelPushConfig based on starting cell address in an Excel-readable string format and workbook properties.")]
        [Input("startingCell", "Starting cell address in an Excel-readable string format.")]
        [InputFromProperty("workbookProperties")]
        [Output("config", "ExcelPushConfig created based on the inputs.")]
        public static ExcelPushConfig ExcelPushConfig(string startingCell = "", WorkbookProperties workbookProperties = null)
        {
            CellAddress topLeft = null;
            if (!string.IsNullOrWhiteSpace(startingCell))
            {
                topLeft = Create.CellAddress(startingCell);
                if (topLeft == null)
                    return null;
            }

            return new ExcelPushConfig { StartingCell = topLeft, WorkbookProperties = workbookProperties };
        }

        /*******************************************/
    }
}

