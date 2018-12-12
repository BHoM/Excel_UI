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

namespace BH.UI.Excel
{
    public static class Properties
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get all properties from an object. WARNING This is an array formula and will take up more than one cell!", Category = "BHoM")]
        public static object Explode(
                [ExcelArgument(Name = "object ids")] object objectIds,
                [ExcelArgument(Name = "Include the name of the properties")] bool includePropertyNames = false,
                [ExcelArgument(Name = "Explode inner objects")] bool goDeep = false)
        {

            object[] _objectIds = new object[] { };
            if (objectIds is object[,])
            {
                 _objectIds = (objectIds as object[,]).Cast<object>().ToArray().CleanArray();
            } else if (objectIds is object[])
            {
                _objectIds = (objectIds as object[]);
            } else if (objectIds is string)
            {
                _objectIds = new[] { objectIds };
            }

            //Clean the array
            _objectIds = _objectIds.CleanArray();

            //Get the object
            List<object> objs = _objectIds.Select(x => {
                string str = x as string;
                int start = str.LastIndexOf("[");
                int end = str.LastIndexOf("]");
                if(start != -1 && end != -1 && end > start)
                {
                    start++;
                    return Project.ActiveProject.GetAny(str.Substring(start, end - start));
                }
                return x;
            }).ToList();

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
                return ArrayResizer.Resize(outArr);
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

                return ArrayResizer.Resize(outArr);
                //return ArrayResizer.Resize( outArr);
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
                if (obj is IEnumerable)
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
