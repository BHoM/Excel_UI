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
using FastMember;
using System.Drawing;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** public Methods                    ****/
        /*******************************************/
        public static BH.oM.Excel.Font Font(IXLFont xLFont)
        {
            return new BH.oM.Excel.Font()
            {
                Bold = xLFont.Bold,
                FontColour = xLFont.FontColor.ColorType == XLColorType.Theme ? Color.Black : xLFont.FontColor.Color,
                FontName = xLFont.FontName,
                FontSize = xLFont.FontSize,
                Italic  = xLFont.Italic,
                Shadow = xLFont.Shadow,
                Strikethrough = xLFont.Strikethrough,
                Underline = (UnderlineStyle)(int)xLFont.Underline
            };
        }

        /*******************************************/
        /**** private Methods                   ****/
        /*******************************************/
        private static Color ToColor(this XLThemeColor xLThemeColor, IXLWorkbook wb)
        {
            string themeColourName = xLThemeColor.ToString();
            XLColor color = ObjectAccessor.Create(wb.Theme)[themeColourName] as XLColor;
            return color.Color;
        }
     }
}
