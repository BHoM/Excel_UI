/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
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
using BH.oM.Base.Attributes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Get all properties from an object. WARNING This is an array formula and will take up more than one cell!")]
        [Input("objects", "Objects to explode")]
        [Input("includePropertyNames", "Include the name of the properties")]
        [Input("goDeep", "Explode inner objects")]
        [Input("transpose", "Transpose the resulting table (i.e. one object per column instead of per row)")]
        public static object Explode(this List<object> objects, bool includePropertyNames = false, bool goDeep = false, bool transpose = false)
        {
            Engine.Base.Compute.ClearCurrentEvents();
            if (objects == null || objects.Count == 0)
                return "No objects to explode";

            // Clean the list
            List<object> objs = objects.FindAll(item => item != null);
            if (objs == null || objs.Count == 0)
                return "Failed to get non null objects";

            // Get the property dictionary for the object
            List<Dictionary<string, object>> props = GetPropertyDictionaries(objs, goDeep);
            if (props.Count < 1)
                return "Failed to get properties";

            // Get the exploded table
            List<List<object>> result = new List<List<object>>();
            List<string> keys = props.SelectMany(x => x.Keys).Distinct().ToList();

            if (includePropertyNames)
                result.Add(keys.ToList<object>());

            for (int i = 0; i < props.Count; i++)
                result.Add(keys.Select(k => props[i].ContainsKey(k) ? props[i][k] : null).ToList());

            if (transpose)
            {
                result = result.SelectMany(row => row.Select((value, index) => new { value, index }))
                    .GroupBy(cell => cell.index, cell => cell.value)
                    .Select(g => g.ToList()).ToList();
            }

            return result;
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static List<Dictionary<string, object>> GetPropertyDictionaries(List<object> objs, bool goDeep = false)
        {
            //Get the property dictionary for the object
            List<Dictionary<string, object>> props = new List<Dictionary<string, object>>();
            foreach (object obj in objs)
            {
                if (obj is IEnumerable && !(obj is string))
                {
                    props.AddRange(GetPropertyDictionaries((obj as IEnumerable).Cast<object>().ToList(), goDeep));
                }
                else
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    GetPropertyDictionary(ref dict, obj, goDeep);
                    props.Add(dict);
                }
            }

            return props;
        }


        /*******************************************/

        private static void GetPropertyDictionary(ref Dictionary<string, object> dict, object obj, bool goDeep = false, string parentType = "")
        {
            if (obj == null)
            {
                return;
            }
            else if (obj.GetType().IsPrimitive || obj is string || obj is Guid || obj is Enum)
            {
                string key = parentType.Length > 0 ? parentType : "Value";
                dict[key] = obj;
                return;
            }
            else
            {
                foreach (KeyValuePair<string, object> kvp in obj.PropertyDictionary())
                {
                    string key = (parentType.Length > 0) ? parentType + "." + kvp.Key : kvp.Key;
                    if (goDeep)
                        GetPropertyDictionary(ref dict, kvp.Value, true, key);
                    else
                        dict[key] = kvp.Value;
                }
            }
        }

        /*******************************************/
    }
}


