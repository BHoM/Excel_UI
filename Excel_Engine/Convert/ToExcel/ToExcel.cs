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

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.Engine.Reflection;
using BH.oM.Reflection.Attributes;
using ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Convert
    {
        /*******************************************/
        /**** Interface Methods                 ****/
        /*******************************************/

        public static object IToExcel(this object item)
        {
            return ToExcel(item as dynamic);
        }


        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        [Description("A DateTime object to an OADate number for MS Office.")]
        [Input("dateTime", "The Date and Time to convert.")]
        [Output("An OADate number.")]
        public static double ToExcel(this DateTime dateTime)
        {
            return dateTime.ToOADate();
        }

        /*******************************************/

        [Description("Converts a BHoM Reference to an ExcelReference object.")]
        [Input("omRef", "The reference to convert.")]
        [Output("An ExcelDNA ExcelReference.")]
        public static ExcelReference ToExcel(this oM.Excel.Reference omRef)
        {
            var rects = omRef.Rectangles.Select((rect) =>
                new int[] { rect.RowFirst, rect.RowLast, rect.ColumnFirst, rect.ColumnLast }).ToArray();
            return new ExcelReference(rects, omRef.Sheet);
        }



        /*******************************************/
        /**** Fallback Methods                  ****/
        /*******************************************/

        private static object ToExcel(object data)
        {
            return data;
        }

        /*******************************************/
    }
}
