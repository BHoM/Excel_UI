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

using BH.Engine.Excel;
using BH.Engine.Reflection;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using ExcelDna.Integration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Callers
{
    class ExplodeCaller : BH.UI.Base.Caller
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override System.Drawing.Bitmap Icon_24x24 { get { return m_Native.Icon_24x24; } }

        public override string Name { get { return m_Native.Name; } }

        public override string Category { get { return m_Native.Category; } }

        public override string Description { get { return m_Native.Description; } }

        public override int GroupIndex { get { return m_Native.GroupIndex; } }


        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ExplodeCaller() : base(typeof(ExplodeCaller).GetMethod("Explode")) {}


        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        [ExcelFunction(Description = "Get all properties from an object. WARNING This is an array formula and will take up more than one cell!", Category = "BHoM")]
        public static object Explode(
                [ExcelArgument(Name = "Objects")] List<object> objects,
                [ExcelArgument(Name = "Include the name of the properties")] bool includePropertyNames = false,
                [ExcelArgument(Name = "Explode inner objects")] bool goDeep = false,
                [ExcelArgument(Name = "Transpose")] bool transpose = false)
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
            object[,] outArr;
            if (includePropertyNames)
            {
                //Create an 2d array to contain property names and values
                outArr = new object[props.Count + 1, props[0].Count];
                int counter = 0;

                foreach (KeyValuePair<string, object> kvp in props[0])
                {
                    outArr[0, counter] = kvp.Key;
                    outArr[1, counter] = FormulaDataAccessor.ToExcel(kvp.Value);
                    counter++;
                }

                for (int i = 1; i < props.Count; i++)
                {
                    counter = 0;
                    foreach (KeyValuePair<string, object> kvp in props[i])
                    {
                        outArr[i + 1, counter] = FormulaDataAccessor.ToExcel(kvp.Value);
                        counter++;
                    }
                }
            }
            else
            {
                //Create an object array to contain the property values
                outArr = new object[props.Count, props[0].Count];

                for (int i = 0; i < props.Count; i++)
                {
                    int counter = 0;
                    foreach (KeyValuePair<string, object> kvp in props[i])
                    {
                        outArr[i, counter] = FormulaDataAccessor.ToExcel(kvp.Value);
                        counter++;
                    }
                }
            }

            if (transpose)
                outArr = Transpose(outArr);

            //Output the values as an array
            return outArr; 
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static object[,] Transpose(object[,] arr)
        {
            int width = arr.GetLength(0);
            int height = arr.GetLength(1);
            object[,] transposed = new object[height, width];
            for (int i = 0; i < width * height; i++)
            {
                int x = i % width;
                int y = i / width;
                transposed[y, x] = arr[x, y];
            }
            return transposed;
        }
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

                foreach (KeyValuePair<string, object> kvp in baseDict)
                {
                    object value = FormulaDataAccessor.ToExcel(kvp.Value);
                    object innerObj = AddIn.GetObject(value.ToString());

                    if (innerObj == null || kvp.Key == "BHoM_Guid")
                        dict[parentType + kvp.Key] = value;
                    else
                    {
                        GetPropertyDictionary(ref dict, innerObj, true, parentType + kvp.Key + ": ");
                    }
                }
            }
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private BH.UI.Base.Components.ExplodeCaller m_Native = new BH.UI.Base.Components.ExplodeCaller();

        /*******************************************/
    }
}
