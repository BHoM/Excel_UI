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

using BH.Engine.Reflection;
using BH.oM.Adapters.Excel;
using BH.oM.Base.Attributes;
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
                BH.Engine.Base.Compute.RecordError("Cell address cannot be null.");
                return false;
            }

            if (address.Row < 1)
            {
                BH.Engine.Base.Compute.RecordError("Row index cannot lower than 1.");
                return false;
            }

            if (!m_ColumnIndexFormat.IsMatch(address.Column))
            {
                BH.Engine.Base.Compute.RecordError($"Column label equal to {address.Column} is invalid, it needs to consist of capital letters only.");
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
                BH.Engine.Base.Compute.RecordError("Cell range cannot be null.");
                return false;
            }

            return range.From.IsValid() && range.To.IsValid();
        }

        /*******************************************/

        [Description("Checks whether given cell address in an Excel-readable string format is valid for use in Excel adapter and raises errors if not.")]
        [Input("address", "Cell address in an Excel-readable string format to be validated.")]
        [Output("valid", "True if the input string is valid, otherwise false.")]
        public static bool IsValidAddress(this string address)
        {
            if (address == null)
            {
                BH.Engine.Base.Compute.RecordError("Cell address cannot be null.");
                return false;
            }

            if (!m_AddressFormat.IsMatch(address))
            {
                BH.Engine.Base.Compute.RecordError($"Address equal to {address} is not valid: it needs to consist of capital letters followed by digits.");
                return false;
            }

            int row = int.Parse(Regex.Match(address, @"\d+").Value);
            if (row < 1)
            {
                BH.Engine.Base.Compute.RecordError($"Address equal to { address} is not valid: row index cannot lower than 1.");
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
                BH.Engine.Base.Compute.RecordError("Cell range cannot be null.");
                return false;
            }

            if (!m_RangeFormat.IsMatch(range))
            {
                BH.Engine.Base.Compute.RecordError($"Range equal to {range} is not valid: it needs to consist of two sets of capital letters followed by digits, divided with a colon, for example 'A1:Z99'.");
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

