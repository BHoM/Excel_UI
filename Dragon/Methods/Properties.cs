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

        [ExcelFunction(Description = "Set the property of an object", Category = "Dragon")]
        public static object SetProperty(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] object[] propNames,
            [ExcelArgument(Name = "property value")] object[] propValues)
        {

            IObject obj = Project.ActiveProject.GetObject(objectId);

            if (obj == null)
                return "Failed to get BHoMObject";

            IObject clone = obj.GetShallowClone(true);

            string message;
            
            if (!InOutHelp.SetPropertyHelper(clone, propNames, propValues, out message))
                return message;

            Project.ActiveProject.AddObject(clone);
            return clone.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Set the property of an object", Category = "Dragon")]
        public static object SetPropertyList(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] object[] propName,
            [ExcelArgument(Name = "property value")] object[] propValues)
        {

            IObject obj = Project.ActiveProject.GetObject(objectId);

            if (obj == null)
                return "Failed to get BHoMObject";

            IObject clone = obj.GetShallowClone(true);

            string message;

            if (!InOutHelp.SetPropertyHelper(clone, propName, propValues, out message))
                return message;

            Project.ActiveProject.AddObject(clone);
            return clone.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Adds a custom data to an object", Category = "Dragon")]
        public static object AddCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key,
            [ExcelArgument(Name = "Custom data value")] object val)
        {
            IObject oblObj = Project.ActiveProject.GetObject(objectId);
            IObject newObj = oblObj.GetShallowClone(true);

            newObj.CustomData[key] = val;

            Project.ActiveProject.AddObject(newObj);
            return newObj.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Gets a custom data from an object", Category = "Dragon")]
        public static object GetCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key)
        {
            IObject obj = Project.ActiveProject.GetObject(objectId);

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

            return XlCall.Excel(XlCall.xlUDF, "Resize", obj.PropertyNames().ToArray());
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get all properties from an object. WARNING This is an array formula and will take up more than one cell!", Category = "Dragon")]
        public static object Explode(
                [ExcelArgument(Name = "object id")] string objectId,
                 [ExcelArgument(Name = "Include the name of the properties")] bool includePropertyNames = false)
        {
            //Get the object
            object obj = Project.ActiveProject.GetAny(objectId);

            if (obj == null)
                return "Failed to get object";

            //Get the property dictionary for the object
            Dictionary<string, object> props;
            if (obj is IExcelObject)
                props = ((IExcelObject)obj).PropertyDictionary();
            else
                props = obj.PropertyDictionary();
            

            if (includePropertyNames)
            {
                //Create an 2d array to contain property names and values
                object[,] outArr = new object[2, props.Count];
                int counter = 0;
                foreach (KeyValuePair<string, object> kvp in props)
                {
                    outArr[0, counter] = kvp.Key;
                    outArr[1, counter] = kvp.Value.ReturnTypeHelper();
                    counter++;
                }

                //Output the values as an array
                return XlCall.Excel(XlCall.xlUDF, "Resize", outArr);
            }
            else
            {
                //Create an object array to contain the property values
                object[] outArr = new object[props.Count];
                int counter = 0;
                foreach (KeyValuePair<string, object> kvp in props)
                {
                    outArr[counter] = kvp.Value.ReturnTypeHelper();
                    counter++;
                }

                return XlCall.Excel(XlCall.xlUDF, "Resize", outArr);
            }
        }

        /*****************************************************************/


        
    }
}
