using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.Engine.Reflection;

namespace BH.UI.Dragon.Base
{
    public static class Base
    {
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM object", Category = "Dragon")]
        public static object CreateObject(
            [ExcelArgument(Name = "object type")] string typeString,
            [ExcelArgument(Name = "property names (optional)")] object[] propNames,
            [ExcelArgument(Name = "property values (optional)")] object[] propValues)
        {

            if (propNames.Length != propValues.Length)
                return "Need to provide the same number of property names as property values";

            Type type = Create.Type(typeString);
            BHoMObject obj = type.GetConstructor(Type.EmptyTypes).Invoke(new object[] { }) as BHoMObject;

            string message;
            if (!SetPropertyHelper(obj, propNames, propValues, out message))
                return message;

            Project.ActiveProject.AddObject(obj);
            return obj.BHoM_Guid.ToString();

            //for (int i = 0; i < propNames.Length; i++)
            //{
            //    if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
            //        continue;

            //    object val;

            //    if (propValues[i] is Guid)
            //    {
            //        val = Project.ActiveProject.GetObject((Guid)propValues[i]);
            //    }
            //    else
            //    {
            //        val = propValues[i];
            //    }

            //    if (!obj.SetPropertyValue(propNames[i] as string, val))
            //        return propNames[i].ToString() + " is not a valid property for the type";
            //}

            //Project.ActiveProject.AddObject(obj);
            //return obj.BHoM_Guid;

            //BHB.BHoMObject newObject = BHB.BHoMObject.FromTypeName(typeString);

            //int nb = Math.Min(propNames.Length, propValues.Length);
            //for (int i = 0; i < nb; i++)
            //{
            //    if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
            //        continue;

            //    System.Reflection.PropertyInfo prop = newObject.GetType().GetProperty(propNames[i] as string);
            //    if (prop == null) continue;

            //    if (prop.PropertyType.IsSubclassOf(typeof(BHB.BHoMObject)))
            //        prop.SetValue(newObject, BHG.Project.ActiveProject.GetObject(propValues[i] as string));
            //    else
            //        prop.SetValue(newObject, propValues[i]);
            //}

            //BHG.Project.ActiveProject.AddObject(newObject);
            //return newObject.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get the property of an object", Category = "Dragon")]
        public static object GetProperty(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] string property)
        {

            object obj = Project.ActiveProject.GetObject(objectId).PropertyValue(property);

            if (obj == null)
                return "Failed to get property";
            else if (obj is BHoMObject)
                return ((BHoMObject)obj).BHoM_Guid.ToString();

            return obj.ToString();


            //BHoMObject obj = Project.ActiveProject.GetObject(objectId);
            //System.Reflection.PropertyInfo propInfo = obj.GetType().GetProperty(property);
            //if (propInfo == null)
            //    return null;

            //object prop = propInfo.GetValue(obj);

            //if (prop is BHB.BHoMObject)
            //    return ((BHB.BHoMObject)prop).BHoM_Guid.ToString();
                
            //return prop.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Set the property of an object", Category = "Dragon")]
        public static object SetProperty(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] object[] propNames,
            [ExcelArgument(Name = "property value")] object[] propValues)
        { 


            BHoMObject obj = Project.ActiveProject.GetObject(objectId);

            if (obj == null)
                return "Failed to get BHoMObject";

            BHoMObject clone = obj.GetShallowClone(true);

            string message;
            if (!SetPropertyHelper(clone, propNames, propValues, out message))
                return message;

            Project.ActiveProject.AddObject(clone);
            return obj.BHoM_Guid;


            //BHB.BHoMObject oldObject = BHG.Project.ActiveProject.GetObject(objectId);

            //BHB.BHoMObject newObject = oldObject.ShallowClone(true);

            //int nb = Math.Min(propNames.Length, propValues.Length);
            //for (int i = 0; i < nb; i++)
            //{
            //    if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
            //        continue;

            //    System.Reflection.PropertyInfo prop = newObject.GetType().GetProperty(propNames[i] as string);
            //    if (prop == null) continue;

            //    if (prop.PropertyType.IsSubclassOf(typeof(BHB.BHoMObject)))
            //        prop.SetValue(newObject, BHG.Project.ActiveProject.GetObject(propValues[i] as string));
            //    else if (prop.PropertyType.IsEnum)
            //        prop.SetValue(newObject, Enum.Parse(prop.PropertyType, propValues[i] as string));
            //    else
            //        prop.SetValue(newObject, propValues[i]);
            //}

            //BHG.Project.ActiveProject.AddObject(newObject);
            //return newObject.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        public static bool SetPropertyHelper(BHoMObject obj, object[] propNames, object[] propValues, out string message)
        {
            message = "";
            for (int i = 0; i < propNames.Length; i++)
            {
                if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
                    continue;

                object val;
                Guid guid;
                if (propValues[i] is string && Guid.TryParse(propValues[i] as string, out guid))
                {
                    val = Project.ActiveProject.GetObject(guid);
                }
                else
                {
                    val = propValues[i];
                }

                if (!obj.SetPropertyValue(propNames[i] as string, val))
                {
                    message = propNames[i].ToString() + " is not a valid property for the type";
                    return false;
                }
            }
            return true;
        }


        [ExcelFunction(Description = "Call the ToString() method from an object", Category = "Dragon")]
        public static object ToString(
            [ExcelArgument(Name = "object id")] string objectId)
        {
            BHoMObject obj = Project.ActiveProject.GetObject(objectId);

            return obj.ToString();
        }

        /*****************************************************************/

        //[ExcelFunction(Description = "Get a definition of all the BhoM objects", Category = "Dragon")]
        //public static object GetAllObjectModels()
        //{

        //    List<string[]> data = new List<string[]>();
        //    foreach (KeyValuePair<string, Type> kvp in  Query.GetBHoMTypeDictionary())
        //    {
        //        if (!kvp.Key.Contains('.')) continue; // Need a better way to access each type only once

        //        string[] trow = new string[3];
        //        trow[0] = kvp.Key;
        //        trow[1] = "";
        //        trow[2] = "";
        //        data.Add(trow);

        //        foreach (PropertyInfo prop in kvp.Value.GetProperties())
        //        {
        //            if (prop.CanRead && prop.CanWrite)
        //            {
        //                string[] row = new string[3];
        //                row[0] = "";
        //                row[1] = prop.Name;
        //                row[2] = prop.PropertyType.ToString();
        //                data.Add(row);
        //            }
        //        }
        //    }

        //    int nb = data.Count;
        //    object[,] array = new object[nb, 3];
        //    for (int i = 0; i < nb; i++)
        //    {
        //        for (int j = 0; j < 3; j++)
        //            array[i, j] = data[i][j];
        //    }

        //    return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        //}

        /*****************************************************************/

        [ExcelFunction(Description = "Adds a custom data to an object", Category = "Dragon")]
        public static object AddCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key,
            [ExcelArgument(Name = "Custom data value")] object val)
        {
            BHoMObject oblObj = Project.ActiveProject.GetObject(objectId);
            BHoMObject newObj = oblObj.GetShallowClone(true);

            newObj.CustomData[key] = val;

            Project.ActiveProject.AddObject(newObj);
            return newObj.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Adds a custom data to an object", Category = "Dragon")]
        public static object GetCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key)
        {
            BHoMObject obj = Project.ActiveProject.GetObject(objectId);

            object val;
            if (!obj.CustomData.TryGetValue(key, out val))
                return null;

            if (val is BHoMObject)
                return ((BHoMObject)val).BHoM_Guid.ToString();

            return val.ToString();
        }

    }
}
