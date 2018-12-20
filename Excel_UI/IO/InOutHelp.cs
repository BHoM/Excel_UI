/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2018, the respective contributors. All rights reserved.
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
using System.Runtime.CompilerServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.Engine.Reflection;

namespace BH.UI.Excel
{
    public static class InOutHelp
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/



        public static object ReturnTypeHelper(this object obj)
        {
            if (obj == null)
                return ExcelError.ExcelErrorNull;
            else if (obj.GetType().IsPrimitive || obj is string)
                return obj;
            else if (obj is Guid)
                return obj.ToString();
            else
                return obj.GetType().ToText() + " [" + Project.ActiveProject.IAdd(obj) + "]";
        }

        /*****************************************************************/

        private static Type iTupleType = Type.GetType("System.ITuple, mscorlib"); //the ITuple interface is "internal" for some reason. getting it out once via reflection to be used for checking in the method above...

        /*****************************************************************/

        public static bool IsNumeric(this object obj)
        {
            if (obj is double)
                return true;
            if (obj is int)
                return true;
            if (obj is float)
                return true;
            if (obj is decimal)
                return true;
            if (obj is byte)
                return true;

            return false;
        }

        /*****************************************************************/

        public static bool IsValidArray(object[] arr)
        {
            if (arr == null)
                return false;

            if (arr.Length < 1)
                return false;

            if (arr.Length == 1 && arr[0] == ExcelMissing.Value)
                return false;

            return true;
        }

        /*****************************************************************/

        public static object[] CleanArray(this object[] arr)
        {
            return arr.Where(x => x != null && x != ExcelMissing.Value && x != ExcelEmpty.Value).ToArray();
        }

        /*****************************************************************/
    }
}
