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

using BH.Engine.Reflection;
using BH.oM.Adapters.Excel;
using BH.oM.Reflection.Attributes;
using System;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Checks whether given BHoM CellAddress is valid for use in Excel adapter and raises errors if not.")]
        [Input("address", "BHoM CellAddress to be validated.")]
        [Output("valid", "True if the input CellAddress is valid, otherwise false.")]
        public static bool IsValid(this CellAddress address)
        {
            if (address == null)
            {
                BH.Engine.Reflection.Compute.RecordError("Cell address cannot be null.");
                return false;
            }

            if (address.Row < 1)
            {
                BH.Engine.Reflection.Compute.RecordError("Row index cannot lower than 1.");
                return false;
            }

            if (!m_ColumnIndexFormat.IsMatch(address.Column))
            {
                BH.Engine.Reflection.Compute.RecordError($"Column label equal to {address.Column} is invalid, it needs to consist of capital letters only.");
                return false;
            }

            return true;
        }

        /*******************************************/

        [Description("Checks whether given BHoM CellRange is valid for use in Excel adapter and raises errors if not.")]
        [Input("range", "BHoM CellRange to be validated.")]
        [Output("valid", "True if the input CellRange is valid, otherwise false.")]
        public static bool IsValid(this CellRange range)
        {
            if (range == null)
            {
                BH.Engine.Reflection.Compute.RecordError("Cell range cannot be null.");
                return false;
            }

            return range.From.IsValid() && range.To.IsValid();
        }

        /*******************************************/

        [Description("Checks whether the given object is a valid label or index (where 1 is equal to 'A') for an Excel column.")]
        [Input("column", "Object to be validated.")]
        [Output("valid", "True if the input object is a valid label or index for an Excel column, otherwise false.")]
        public static bool IsValidColumn(this object column)
        {
            double doubleIndex = 0;

            if (column?.GetType() == typeof(string))
            {
                // If the input is a string, check if we can parse it as a number, or if it matches the Excel column label format.
                if (!double.TryParse(column.ToString(), out doubleIndex) && !m_ColumnIndexFormat.IsMatch(column.ToString()))
                {
                    BH.Engine.Reflection.Compute.RecordError($"Column label `{column.ToString()}` is invalid. Make sure it consists of capital letters only." +
                        $"\nEither specify a text with the Excel column name (e.g. 'AA') or an integer > 1 indicating the column index.");
                    return false;
                }
                else
                    return true;
            }
            else
                // If the input is not a string, try parsing it to a number.
                double.TryParse(column.ToString(), out doubleIndex);


            if (doubleIndex % 1 != 0) // it's not an integer.
            {
                BH.Engine.Reflection.Compute.RecordError($"Input `{column.ToString()}` can not be converted to an integer value indicating a column." +
                    $"\nEither specify a text with the Excel column name (e.g. 'AA') or an integer > 1 indicating the column index.");

                return false;
            }

            if (doubleIndex < 1)
            {
                BH.Engine.Reflection.Compute.RecordError("Index smaller than 1 is not allowed. 1 corresponds to Excel column 'A'.");
                return false;
            }

            return true;
        }

        /*******************************************/

        [Description("Checks whether the given object is a valid label for an Excel row.")]
        [Input("row", "Object to be validated.")]
        [Output("valid", "True if the input object is a valid label for an Excel row, otherwise false.")]
        public static bool IsValidRow(this object row)
        {
            double doubleIndex = 0;

            if (!double.TryParse(row.ToString(), out doubleIndex))
            {
                BH.Engine.Reflection.Compute.RecordError($"Row label `{row.ToString()}` is invalid." +
                    $"The {nameof(row)} input must either be integer number or a string integer number indicating the row index.");

                return false;
            }


            if (doubleIndex % 1 != 0) // it's not an integer.
            {
                BH.Engine.Reflection.Compute.RecordError($"Input `{row.ToString()}` can not be converted to an integer value indicating a row." +
                    $"The {nameof(row)} input must either be integer number or a string integer number indicating the row index.");

                return false;
            }

            if (doubleIndex < 1)
            {
                BH.Engine.Reflection.Compute.RecordError("Index smaller than 1 is not allowed. Index 1 corresponds to the first Excel row.");
                return false;
            }

            return true;
        }

        /*******************************************/

        [Description("Checks whether given cell address in an Excel-readable string format is valid for use in Excel adapter and raises errors if not.")]
        [Input("address", "Cell address in an Excel-readable string format to be validated.")]
        [Output("valid", "True if the input string is valid, otherwise false.")]
        public static bool IsValidAddress(this string address)
        {
            if (address == null)
            {
                BH.Engine.Reflection.Compute.RecordError("Cell address cannot be null.");
                return false;
            }

            if (!m_AddressFormat.IsMatch(address))
            {
                BH.Engine.Reflection.Compute.RecordError($"Address equal to {address} is not valid: it needs to consist of capital letters followed by digits.");
                return false;
            }

            int row = int.Parse(Regex.Match(address, @"\d+").Value);
            if (row < 1)
            {
                BH.Engine.Reflection.Compute.RecordError($"Address equal to { address} is not valid: row index cannot lower than 1.");
                return false;
            }

            return true;
        }

        /*******************************************/

        [Description("Checks whether given cell range in an Excel-readable string format is valid for use in Excel adapter and raises errors if not.")]
        [Input("range", "Cell range in an Excel-readable string format to be validated.")]
        [Output("valid", "True if the input string is valid, otherwise false.")]
        public static bool IsValidRange(this string range)
        {
            if (range == null)
            {
                BH.Engine.Reflection.Compute.RecordError("Cell range cannot be null.");
                return false;
            }

            if (!m_RangeFormat.IsMatch(range))
            {
                BH.Engine.Reflection.Compute.RecordError($"Range equal to {range} is not valid: it needs to consist of two sets of capital letters followed by digits, divided with a colon, for example 'A1:Z99'.");
                return false;
            }

            string[] split = range.Split(new char[] { ':' });
            string from = split[0];
            string to = split[1];

            return from.IsValidAddress() && to.IsValidAddress();
        }


        /*******************************************/
        /**** Private fields                    ****/
        /*******************************************/

        private static readonly Regex m_ColumnIndexFormat = new Regex(@"^[A-Z]+$");
        private static readonly Regex m_AddressFormat = new Regex(@"^[A-Z]+\d+$");
        private static readonly Regex m_RangeFormat = new Regex(@"^[A-Z]+\d+:[A-Z]+\d+$");

        /*******************************************/
    }
}
