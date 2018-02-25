using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
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
    public static class InOutHelp
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        public static object CheckAndGetStoredObject(this object obj)
        {
            Guid guid;
            if (obj is string && Guid.TryParse(obj as string, out guid))
            {
                //Get out object or geometry
                return Project.ActiveProject.GetAny(guid);
            }
            else
            {
                return obj;
            }
        }

        /*****************************************************************/
        
        public static bool SetPropertyHelper(this object obj, object[] propNames, object[] propValues, out string message)
        {
            message = "";
            for (int i = 0; i < propNames.Length; i++)
            {
                if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
                    continue;

                object val = propValues[i].CheckAndGetStoredObject();
                val = val is IExcelObject ? ((IExcelObject)val).InnerObject : val;

                //Set the properties
                if (!obj.SetPropertyValue(propNames[i] as string, val))
                {
                    message = propNames[i].ToString() + " is not a valid property for the type";
                    return false;
                }
            }
            return true;
        }

        /*****************************************************************/

        public static bool SetPropertyHelper(this CustomObject obj, object[] propNames, object[] propValues, out string message)
        {
            message = "";
            for (int i = 0; i < propNames.Length; i++)
            {
                if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
                    continue;

                object val = propValues[i].CheckAndGetStoredObject();

                val = val is IExcelObject ? ((IExcelObject)val).InnerObject : val;

                //Set the properties
                if (propNames[i] as string == "Name")
                {
                    if (!obj.SetPropertyValue(propNames[i] as string, val))
                    {
                        message = propNames[i].ToString() + " is not a valid property for the type";
                        return false;
                    }
                }
                else
                {
                    obj.CustomData[propNames[i] as string] = val;
                }
            }
            return true;
        }


        /*****************************************************************/

        //public static bool SetPropertyHelper(this IGeometry obj, object[] propNames, object[] propValues, out string message)
        //{
        //    message = "";
        //    for (int i = 0; i < propNames.Length; i++)
        //    {
        //        if ((propNames[i] is ExcelMissing) || (propValues[i] is ExcelMissing))
        //            continue;

        //        object val = propValues[i].CheckAndGetObjectOrGeometry();

        //        if (!obj.SetPropertyValue(propNames[i] as string, val))
        //        {
        //            message = propNames[i].ToString() + " is not a valid property for the type";
        //            return false;
        //        }
        //    }
        //    return true;
        //}

        /*****************************************************************/

        public static object ReturnTypeHelper(this object obj)
        {
            if (obj == null)
                return "Failed to get property";
            else if (IsNumeric(obj))
                return obj;
            else if (obj is IBHoMObject)
            {
                IBHoMObject iObj = (IBHoMObject)obj;
                Project.ActiveProject.Add(iObj);
                return iObj.BHoM_Guid.ToString();
            }
            else if (obj is IGeometry)
            {
                return Project.ActiveProject.Add(obj as IGeometry).ToString();
            }
            else if (obj is IDictionary)
            {

                //Special case for the dictionary of <string,object> to avoid using reflection for this common type
                if (obj is Dictionary<string, object>)
                    return Project.ActiveProject.Add(new ExcelDictionary<string, object>() { Data = (Dictionary<string, object>)obj }).ToString();

                //Use reflection to instansiate any other type of dictionary
                Type type = typeof(ExcelDictionary<,>).MakeGenericType(obj.GetType().GetGenericArguments());
                var prop = type.GetProperty("Data");

                var dict = Activator.CreateInstance(type);
                prop.SetValue(dict, obj);

                return Project.ActiveProject.IAdd(dict).ToString();

            }
            else if (obj is IList)
            {
                Type type = typeof(ExcelList<>).MakeGenericType(obj.GetType().GetGenericArguments());
                var prop = type.GetProperty("Data");

                var list = Activator.CreateInstance(type);
                prop.SetValue(list, obj);
                return Project.ActiveProject.IAdd(list).ToString();
            }
            else if (iTupleType.IsAssignableFrom(obj.GetType()))
            {
                Type type = typeof(ExcelTuple<,>).MakeGenericType(obj.GetType().GetGenericArguments());
                var prop = type.GetProperty("Data");

                var tuple = Activator.CreateInstance(type);
                prop.SetValue(tuple, obj);
                return Project.ActiveProject.IAdd(tuple).ToString();
            }


            return obj.ToString();
        }

        /*****************************************************************/

        private static Type iTupleType = Type.GetType("System.ITuple, mscorlib"); //the ITuple interface is "internal" for some reason. getting it out once via reflection to be used for checking in the method above...

        /*****************************************************************/

        public static bool IsNumeric(this object obj)
        {
            if (obj is double)
                return true;
            if (obj is int)
                return true;
            if (obj is float)
                return true;
            if (obj is decimal)
                return true;
            if (obj is byte)
                return true;

            return false;
        }

        /*****************************************************************/

        public static bool IsValidArray(object[] arr)
        {
            if (arr == null)
                return false;

            if (arr.Length < 1)
                return false;

            if (arr.Length == 1 && arr[0] == ExcelMissing.Value)
                return false;

            return true;
        }

        /*****************************************************************/
    }
}
