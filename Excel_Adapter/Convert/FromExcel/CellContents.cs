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
using ClosedXML.Excel;
using System;
using System.ComponentModel;

namespace BH.Adapter.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Converts the given ClosedXML cell contents object to a BHoM CellContents.")]
        [Input("xLCell", "ClosedXML cell contents object to convert from.")]
        [Output("cellContents", "BHoM CellContents based on the input ClosedXML cell contents object.")]
        public static CellContents FromExcel(this IXLCell xLCell)
        {
            if (xLCell == null)
                return null;

            return new CellContents()
            {
                Comment = xLCell.HasComment ? xLCell.Comment.Text : "",
                Value = xLCell.Value,
                Address = BH.Engine.Excel.Create.CellAddress(xLCell.Address.ToString()),
                DataType = xLCell.DataType.SystemType(),
                FormulaA1 = xLCell.FormulaA1,
                FormulaR1C1 = xLCell.FormulaR1C1,
                HyperLink = xLCell.HasHyperlink ? xLCell.Hyperlink.ExternalAddress.ToString() : "",
                RichText = xLCell.HasRichText ? xLCell.RichText.Text : ""
            };
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static Type SystemType(this XLDataType dataType)
        {
            switch (dataType)
            {
                case XLDataType.Boolean:
                    return typeof(bool);
                case XLDataType.DateTime:
                    return typeof(DateTime);
                case XLDataType.Number:
                    return typeof(double);
                case XLDataType.Text:
                    return typeof(string);
                case XLDataType.TimeSpan:
                    return typeof(TimeSpan);
                default:
                    return null;
            }
        }

        /*******************************************/
    }
}

