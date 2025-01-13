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

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.oM.Base;
using BH.Engine.Base;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** FromExcel Methods                 ****/
        /*******************************************/

        public static object FromExcel(object input)
        {
            if (input == null)
                return null;
            else if (input.GetType().IsPrimitive)
                return input;
            else if (input is string)
            {
                object obj = AddIn.GetObject(input as string);
                return obj == null ? input : obj;
            }
            else if (input is object[,])
                return FromExcel(input as object[,]);
            else
                return input;
        }

        /*******************************************/

        public static object[] FromExcel(object[] input)
        {
            return input.Select(x => FromExcel(x)).ToArray();
        }

        /*******************************************/

        public static object[,] FromExcel(object[,] input)
        {
            int height = input.GetLength(0);
            int width = input.GetLength(1);

            object[,] evaluated = new object[height, width];
            for (int i = 0; i < width; i++)
            {
                for (int j = 0; j < height; j++)
                    evaluated[j, i] = FromExcel(input[j, i] is ExcelEmpty ? null : input[j, i]);
            }
            return evaluated;
        }


        /*******************************************/
        /**** ToExcel Methods                   ****/
        /*******************************************/

        public static object ToExcel(object data) 
        {
            try
            {
                if (data == null)
                    return ExcelError.ExcelErrorNull;
                else if (data.GetType().IsPrimitive || data is string || data is object[,])
                    return data;
                else if (data is Guid)
                    return data.ToString();
                else if (data is IEnumerable && !(data is ICollection))
                    return ToExcel((data as IEnumerable).Cast<object>().ToList());
                else if (data.GetType().IsEnum)
                    return System.Enum.GetName(data.GetType(), data);
                else if (data is DateTime)
                {
                    DateTime? date = data as DateTime?;
                    if (date.HasValue)
                        return date.Value.ToOADate();
                }

                string name = "";
                if (data is Type)
                    name = ((Type)data).ToText(true);
                else
                    name = data.GetType().ToText();

                return name + " [" + AddIn.IAddObject(data) + "]";
            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }
        }

        /*******************************************/

        public static object[] ToExcel(object[] input)
        {
            if (input == null)
                return new object[] { ExcelError.ExcelErrorNull };

            return input.Select(x => ToExcel(x)).ToArray();
        }

        /*******************************************/

        public static object[,] ToExcel(object[,] input)
        {
            if (input == null)
                return new object[,] { { ExcelError.ExcelErrorNull } };

            int height = input.GetLength(0);
            int width = input.GetLength(1);

            object[,] evaluated = new object[height, width];
            for (int i = 0; i < width; i++)
            {
                for (int j = 0; j < height; j++)
                    evaluated[j, i] = ToExcel(input[j, i]);
            }
            return evaluated;
        }

        /*******************************************/

        public static object[,] ToExcel(List<List<object>> input)
        {
            if (input == null)
                return new object[,] { { ExcelError.ExcelErrorNull } };

            int height = input.Count;
            int width = input.Select(x => x.Count).Min();

            object[,] evaluated = new object[height, width];
            for (int i = 0; i < width; i++)
            {
                for (int j = 0; j < height; j++)
                    evaluated[j, i] = ToExcel(input[j][i]);
            }
            return evaluated;
        }

        /*******************************************/
    }
}






