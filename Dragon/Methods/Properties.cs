using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.Engine.Reflection;
using BH.oM.Geometry;

namespace BH.UI.Dragon
{
    public static class Properties
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get the property of an object", Category = "Dragon")]
        public static object GetProperty(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] string property)
        {

            //Get out the object
            object obj = Project.ActiveProject.GetAny(objectId);

            if (obj == null)
                return "Object does not exist";

            object prop;
            //Get out the property. If object is custom object look in the custom data dictionary
            if (obj is CustomObject && property != "Name")
                prop = (obj as CustomObject).CustomData[property];    
            else    
                prop = obj.PropertyValue(property);

            return prop.ReturnTypeHelper();

        }

        /*****************************************************************/

        [ExcelFunction(Description = "Adds a custom data to an object", Category = "Dragon")]
        public static object AddCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key,
            [ExcelArgument(Name = "Custom data value")] object val)
        {
            IBHoMObject oblObj = Project.ActiveProject.GetBHoM(objectId);
            IBHoMObject newObj = oblObj.GetShallowClone(true);

            newObj.CustomData[key] = val;

            Project.ActiveProject.Add(newObj);
            return newObj.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Gets a custom data from an object", Category = "Dragon")]
        public static object GetCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key)
        {
            IBHoMObject obj = Project.ActiveProject.GetBHoM(objectId);

            object val;
            if (!obj.CustomData.TryGetValue(key, out val))
                return "Custom data with key: " + key + "Does not extist in the custom data";

            return val.ReturnTypeHelper();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Gets all property names from an object. WARNING This is an array formula and will take up more than one cell!", Category = "Dragon")]
        public static object GetAllPropertyNames(
                [ExcelArgument(Name = "object id")] string objectId)
        {
            object obj = Project.ActiveProject.GetAny(objectId);

            return ArrayResizer.Resize( obj.PropertyNames().ToArray());
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get all properties from an object. WARNING This is an array formula and will take up more than one cell!", Category = "Dragon")]
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
            List<object> objs = _objectIds.Select(x => Project.ActiveProject.GetAny(x as string)).ToList();

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
                Dictionary<string, object> dict = new Dictionary<string, object>();
                GetPropertyDictionary(ref dict, obj, goDeep);

                props.Add(dict);
            }

            return props;
        }


        /*****************************************************************/

        private static void GetPropertyDictionary(ref Dictionary<string,object> dict, object obj, bool goDeep = false, string parentType = "")
        {
            if (!goDeep)
            {
                if (obj is IExcelObject)
                    dict = ((IExcelObject)obj).PropertyDictionary();
                else
                    dict = obj.PropertyDictionary();
                return;
            }
            else
            {
                Dictionary<string, object> baseDict;

                if (obj is IExcelObject)
                    baseDict = ((IExcelObject)obj).PropertyDictionary();
                else
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
