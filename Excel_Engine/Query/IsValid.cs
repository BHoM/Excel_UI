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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        //[Description("Get all properties from an object. WARNING This is an array formula and will take up more than one cell!")]
        //[Input("objects", "Objects to explode")]
        //[Input("includePropertyNames", "Include the name of the properties")]
        //[Input("goDeep", "Explode inner objects")]
        //[Input("transpose", "Transpose the resulting table (i.e. one object per column instead of per row)")]
        public static bool IsValid(this CellAddress address)
        {
            if (address == null)
            {
                BH.Engine.Reflection.Compute.RecordError("Cell address cannot be null.");
                return false;
            }

            if (address.RowIndex < 1)
            {
                BH.Engine.Reflection.Compute.RecordError("Row index cannot lower than 1.");
                return false;
            }

            if (!m_ColumnIndexFormat.IsMatch(address.ColumnIndex))
            {
                BH.Engine.Reflection.Compute.RecordError($"Column index equal to {address.ColumnIndex} is invalid, it needs to consist of capital letters only.");
                return false;
            }

            return true;
        }

        /*******************************************/

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
