/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2020, the respective contributors. All rights reserved.
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

using BH.oM.Excel;
using ClosedXML.Excel;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/
        public static CellContents CellContents(IXLCell xLCell)
        {
            return new CellContents()
            {
                Comment = xLCell.HasComment ? xLCell.Comment.Text : "",
                Value = xLCell.Value,
                Address = xLCell.Address.ToString(),
                DataType = xLCell.DataType.GetType(),
                FormulaA1 = xLCell.FormulaA1,
                FormulaR1C1 = xLCell.FormulaR1C1,
                HyperLink = xLCell.HasHyperlink ? xLCell.Hyperlink.ExternalAddress.ToString() : "",
                RichText = xLCell.HasRichText ? xLCell.RichText.Text : ""
            };
        }
    }
}