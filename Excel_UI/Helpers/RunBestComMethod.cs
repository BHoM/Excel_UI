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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using BH.Engine.Reflection;
using System.Reflection;
using System.ComponentModel;

namespace BH.UI.Excel
{
    public static partial class Helpers
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public static object RunBestComMethod<T>(IEnumerable<T> methods, Collection inputs = null) where T : MethodBase
        {
            if (inputs == null)
                inputs = new Collection();
            object[] arguments = inputs.FromCom().Select(x => FixType(x)).ToArray();

            T method = null;
            foreach (T m in methods)
            {
                Type[] paramTypes = m.GetParameters().Select(x => x.ParameterType).ToArray();
                if (paramTypes.Length >= arguments.Length)
                {
                    bool match = true;
                    for (int i = 0; i < arguments.Length; i++)
                        match &= IsAssignableFrom(paramTypes[i], arguments[i]?.GetType());

                    if (match)
                    {
                        method = m;
                        break;
                    }
                }
            }

            if (method == null)
                return null;

            ParameterInfo[] parameters = method.GetParameters();
            if (parameters.Length > arguments.Length)
                arguments = arguments.Concat(parameters.Skip(arguments.Length).Select(x => x.DefaultValue)).ToArray();

            object result = null;
            if (method is ConstructorInfo)
                result = (method as ConstructorInfo).Invoke(arguments);
            else
                result = method.Invoke(null, arguments);

            if (result == null)
                return null;
            else
                return result.IToCom();
        }

        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        private static bool IsAssignableFrom(Type a, Type b)
        {
            if (b == null)
                return true;
            else if (a == null)
                return false;
            else if (a == b)
                return true;
            else if (b.IsPrimitive)
                return TypeDescriptor.GetConverter(b).CanConvertTo(a);
            else
                return a.IsAssignableFrom(b);
        }

        /***************************************************/

        private static object FixType(object argument)
        {
            if (argument is List<object>)
            {
                List<object> list = argument as List<object>;
                List<Type> types = list.Select(x => x.GetType()).Distinct().ToList();
                if (types.Count == 1)
                    return CastList(list, list.First() as dynamic);
                else
                    return argument;
            }
            else
                return argument;
        }

        /***************************************************/

        private static List<T> CastList<T>(List<object> list, T first)
        {
            return list.Cast<T>().ToList();
        }

        /***************************************************/
    }

}

