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
    public static class InOutHelp
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        public static object CheckAndGetObjectOrGeometry(this object obj)
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

        public static object CheckAndGetObject(this object obj)
        {
            Guid guid;
            if (obj is string && Guid.TryParse(obj as string, out guid))
            {
                //Get out object or geometry
                return Project.ActiveProject.GetObject(guid);
            }
            else
            {
                return obj;
            }
        }

        /*****************************************************************/

        public static object CheckAndGetGeometry(this object obj)
        {
            Guid guid;
            if (obj is string && Guid.TryParse(obj as string, out guid))
            {
                //Get out object or geometry
                return Project.ActiveProject.GetGeometry(guid);
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

                object val = propValues[i].CheckAndGetObjectOrGeometry();
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

                object val = propValues[i].CheckAndGetObjectOrGeometry();

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

        //public static bool SetPropertyHelper(this IBHoMGeometry obj, object[] propNames, object[] propValues, out string message)
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
            else if (obj is IObject)
            {
                IObject iObj = (IObject)obj;
                Project.ActiveProject.AddObject(iObj);
                return iObj.BHoM_Guid.ToString();
            }
            else if (obj is IBHoMGeometry)
            {
                return Project.ActiveProject.AddGeometry(obj as IBHoMGeometry).ToString();
            }
            else if (IsNumeric(obj))
                return obj;

            return obj.ToString();
        }

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
