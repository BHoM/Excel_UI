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

using BH.oM.Base;
using System;
using System.ComponentModel;

namespace BH.oM.Adapters.Excel
{
    [Description("Object representing the information stored within a cell: the value and metadata related to it.")]
    public class CellContents : BHoMObject
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        [Description("Comment applied to the cell.")]
        public virtual string Comment { get; set; } = "";

        [Description("Value stored in the cell.")]
        public virtual object Value { get; set; } = null;

        [Description("Address of the cell.")]
        public virtual CellAddress Address { get; set; } = null;

        [Description("Data type of the value stored in the cell. Only 5 data types are considered: number, text, Boolean, date/time, and timespan.")]
        public virtual Type DataType { get; set; } = null;

        [Description("Formula stored in the cell, in standard Excel format (e.g. \"=A1\").")]
        public virtual string FormulaA1 { get; set; } = "";

        [Description("Formula stored in the cell, in R1C1 (relative) format. For more information on that format, please search for 'A1 vs R1C1 Notation'.")]
        public virtual string FormulaR1C1 { get; set; } = "";

        [Description("Hyperlink stored in the cell.")]
        public virtual string HyperLink { get; set; } = "";

        [Description("Information about rich formatting of the cell content.")]
        public virtual string RichText { get; set; } = "";

        /*******************************************/
    }
}


