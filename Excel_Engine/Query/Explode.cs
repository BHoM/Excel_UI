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
            if (obj.GetType().IsPrimitive || obj is string || obj is Guid || obj is Enum)
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
