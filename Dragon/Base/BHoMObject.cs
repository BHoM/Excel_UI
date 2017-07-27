﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using BHB = BHoM.Base;
using BHG = BHoM.Global;
using System.Reflection;

namespace Dragon.Base
{
    public static class BHoMObject
    {
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM object", Category = "Dragon")]
        public static object CreateObject(
            [ExcelArgument(Name = "object type")] string typeString,
            [ExcelArgument(Name = "property names (optional)")] object[] propNames,
            [ExcelArgument(Name = "property values (optional)")] object[] propValues)
        {
            BHB.BHoMObject newObject = BHB.BHoMObject.FromTypeName(typeString);

            int nb = Math.Min(propNames.Length, propValues.Length);
            for (int i = 0; i < nb; i++)
            {
                if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
                    continue;

                System.Reflection.PropertyInfo prop = newObject.GetType().GetProperty(propNames[i] as string);
                if (prop == null) continue;

                if (prop.PropertyType.IsSubclassOf(typeof(BHB.BHoMObject)))
                    prop.SetValue(newObject, BHG.Project.ActiveProject.GetObject(propValues[i] as string));
                else
                    prop.SetValue(newObject, propValues[i]);
            }

            BHG.Project.ActiveProject.AddObject(newObject);
            return newObject.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get the property of an object", Category = "Dragon")]
        public static object GetProperty(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] string property)
        {
            BHB.BHoMObject obj = BHG.Project.ActiveProject.GetObject(objectId);
            System.Reflection.PropertyInfo propInfo = obj.GetType().GetProperty(property);
            if (propInfo == null)
                return null;

            object prop = propInfo.GetValue(obj);

            if (prop is BHB.BHoMObject)
                return ((BHB.BHoMObject)prop).BHoM_Guid.ToString();
                
            return prop.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Set the property of an object", Category = "Dragon")]
        public static object SetProperty(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "property name")] object[] propNames,
            [ExcelArgument(Name = "property value")] object[] propValues)
        {
            BHB.BHoMObject oldObject = BHG.Project.ActiveProject.GetObject(objectId);

            BHB.BHoMObject newObject = oldObject.ShallowClone(true);

            int nb = Math.Min(propNames.Length, propValues.Length);
            for (int i = 0; i < nb; i++)
            {
                if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
                    continue;

                System.Reflection.PropertyInfo prop = newObject.GetType().GetProperty(propNames[i] as string);
                if (prop == null) continue;

                if (prop.PropertyType.IsSubclassOf(typeof(BHB.BHoMObject)))
                    prop.SetValue(newObject, BHG.Project.ActiveProject.GetObject(propValues[i] as string));
                else if (prop.PropertyType.IsEnum)
                    prop.SetValue(newObject, Enum.Parse(prop.PropertyType, propValues[i] as string));
                else
                    prop.SetValue(newObject, propValues[i]);
            }

            BHG.Project.ActiveProject.AddObject(newObject);
            return newObject.BHoM_Guid.ToString();
        }

        /*****************************************************************/


        [ExcelFunction(Description = "Call the ToString() method from an object", Category = "Dragon")]
        public static object ToString(
            [ExcelArgument(Name = "object id")] string objectId)
        {
            BHB.BHoMObject obj = BHG.Project.ActiveProject.GetObject(objectId);

            return obj.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get a definition of all the BhoM objects", Category = "Dragon")]
        public static object GetAllObjectModels()
        {

            List<string[]> data = new List<string[]>();
            foreach (KeyValuePair<string, Type> kvp in  BHB.BHoMObject.TypeDictionary)
            {
                if (!kvp.Key.Contains('.')) continue; // Need a better way to access each type only once

                string[] trow = new string[3];
                trow[0] = kvp.Key;
                trow[1] = "";
                trow[2] = "";
                data.Add(trow);

                foreach (PropertyInfo prop in kvp.Value.GetProperties())
                {
                    if (prop.CanRead && prop.CanWrite)
                    {
                        string[] row = new string[3];
                        row[0] = "";
                        row[1] = prop.Name;
                        row[2] = prop.PropertyType.ToString();
                        data.Add(row);
                    }
                }
            }

            int nb = data.Count;
            object[,] array = new object[nb, 3];
            for (int i = 0; i < nb; i++)
            {
                for (int j = 0; j < 3; j++)
                    array[i, j] = data[i][j];
            }

            return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Adds a custom data to an object", Category = "Dragon")]
        public static object AddCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key,
            [ExcelArgument(Name = "Custom data value")] object val)
        {
            BHB.BHoMObject oblObj = BHG.Project.ActiveProject.GetObject(objectId);
            BHB.BHoMObject newObj = oblObj.ShallowClone(true);

            newObj.CustomData[key] = val;

            BHG.Project.ActiveProject.AddObject(newObj);
            return newObj.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Adds a custom data to an object", Category = "Dragon")]
        public static object GetCustomData(
            [ExcelArgument(Name = "object id")] string objectId,
            [ExcelArgument(Name = "Custom data key")] string key)
        {
            BHB.BHoMObject obj = BHG.Project.ActiveProject.GetObject(objectId);

            object val;
            if (!obj.CustomData.TryGetValue(key, out val))
                return null;

            if (val is BHB.BHoMObject)
                return ((BHB.BHoMObject)val).BHoM_Guid.ToString();

            return val.ToString();
        }

    }
}
