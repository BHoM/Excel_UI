using BH.Engine.Reflection;
using BH.oM.Reflection.Attributes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel
{
    public static class Query
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        [Description("Get all properties from an object. WARNING This is an array formula and will take up more than one cell!")]
        [Input("objects", "Objects to explode")]
        [Input("includePropertyNames", "Include the name of the properties")] 
        [Input("goDeep", "Explode inner objects")] 
        [Input("transpose", "Transpose the resulting table (i.e. one object per column instead of per row)")]
        public static object Explode(List<object> objects, bool includePropertyNames = false, bool goDeep = false, bool transpose = false)
        {
            Engine.Reflection.Compute.ClearCurrentEvents();

            // Clean the list
            List<object> objs = objects.FindAll(item => item != null);

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
                    outArr[1, counter] = kvp.Value;
                    counter++;
                }

                for (int i = 1; i < props.Count; i++)
                {
                    counter = 0;
                    foreach (KeyValuePair<string, object> kvp in props[i])
                    {
                        outArr[i + 1, counter] = kvp.Value;
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
                        outArr[i, counter] = kvp.Value;
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
                Dictionary<string, object> baseDict = obj.PropertyDictionary();
                foreach (KeyValuePair<string, object> kvp in baseDict)
                {
                    if (kvp.Key == "BHoM_Guid")
                        dict[parentType + kvp.Key] = kvp.Value;
                    else
                        GetPropertyDictionary(ref dict, kvp.Value, true, parentType + kvp.Key + ": ");
                }
            }
        }

        /*******************************************/
    }
}
