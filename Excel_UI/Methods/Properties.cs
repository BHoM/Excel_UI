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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.Engine.Reflection;
using System.Collections;
using BH.Engine.Excel;
using BH.UI.Global;

namespace BH.UI.Excel.Methods
{
    public static class Properties
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get all properties from an object. WARNING This is an array formula and will take up more than one cell!", Category = "BHoM")]
        public static object Explode(
                [ExcelArgument(Name = "object ids")] List<object> objects,
                [ExcelArgument(Name = "Include the name of the properties")] bool includePropertyNames = false,
                [ExcelArgument(Name = "Explode inner objects")] bool goDeep = false)
        {
            Engine.Reflection.Compute.ClearCurrentEvents();

            // Clean the list
            List<object> objs = objects.CleanList();

            if (objs == null)
                return "Failed to get object";

            //Get the property dictionary for the object
            List<Dictionary<string, object>> props = GetPropertyDictionaries(objs, goDeep);

            if (props.Count < 1)
                return "Failed to get properties";

            if (includePropertyNames)
            {
                //Create an 2d array to contain property names and values
                object[,] outArr = new object[props.Count +1 , props[0].Count];
                int counter = 0;

                foreach (KeyValuePair<string, object> kvp in props[0])
                {
                    outArr[0, counter] = kvp.Key;
                    outArr[1, counter] = kvp.Value.ReturnTypeHelper();
                    counter++;
                }

                for (int i = 1; i < props.Count; i++)
                {
                    counter = 0;
                    foreach (KeyValuePair<string, object> kvp in props[i])
                    {
                        outArr[i+1, counter] = kvp.Value.ReturnTypeHelper();
                        counter++;
                    }
                }

                //Output the values as an array
                return outArr;
                //return ArrayResizer.Resize( outArr);
            }
            else
            {
                //Create an object array to contain the property values
                object[,] outArr = new object[props.Count, props[0].Count];


                for (int i = 0; i < props.Count; i++)
                {
                    int counter = 0;
                    foreach (KeyValuePair<string, object> kvp in props[i])
                    {
                        outArr[i, counter] = kvp.Value.ReturnTypeHelper();
                        counter++;
                    }
                }

                return outArr;
            }
        }

        /*****************************************************************/
        /******* Private methods                            **************/
        /*****************************************************************/

        private static List<Dictionary<string, object>> GetPropertyDictionaries(List<object> objs, bool goDeep = false)
        {
            //Get the property dictionary for the object
            List<Dictionary<string, object>> props = new List<Dictionary<string, object>>();
            foreach (object obj in objs)
            {
                if (obj is IEnumerable && ! (obj is string) )
                {
                    props.AddRange(GetPropertyDictionaries((obj as IEnumerable).Cast<object>().ToList(), goDeep));
                } else
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    GetPropertyDictionary(ref dict, obj, goDeep);
                    props.Add(dict);
                }

            }

            return props;
        }


        /*****************************************************************/

        private static void GetPropertyDictionary(ref Dictionary<string,object> dict, object obj, bool goDeep = false, string parentType = "")
        {
            if (obj.GetType().IsPrimitive || obj is string)
            {
                dict = new Dictionary<string, object> { { "Value", obj } };
                return;
            }
            if (!goDeep)
            {
                dict = obj.PropertyDictionary();
                return;
            }
            else
            {
                Dictionary<string, object> baseDict;

                baseDict = obj.PropertyDictionary();

                foreach (KeyValuePair<string,object> kvp in baseDict)
                {
                    object value = kvp.Value.ReturnTypeHelper();
                    object innerObj = Project.ActiveProject.GetAny(value.ToString());

                    if (innerObj == null || kvp.Key == "BHoM_Guid")
                        dict[parentType + kvp.Key] = value;
                    else
                    {
                        GetPropertyDictionary(ref dict, innerObj, true, parentType + kvp.Key + ": ");
                    }
                }
            }
        }

        /*****************************************************************/

    }
}

