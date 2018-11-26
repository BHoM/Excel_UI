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
using BH.Engine.Reflection.Convert;

namespace BH.UI.Dragon
{
    public static class InOutHelp
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/



        public static object ReturnTypeHelper(this object obj)
        {
            if (obj == null)
                return ExcelError.ExcelErrorNull;
            else if (obj.GetType().IsPrimitive || obj is string)
                return obj;
            else if (obj is Guid)
                return obj.ToString();
            else
                return obj.GetType().ToText() + " [" + Project.ActiveProject.IAdd(obj) + "]";
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

        public static object[] CleanArray(this object[] arr)
        {
            return arr.Where(x => x != null && x != ExcelMissing.Value && x != ExcelEmpty.Value).ToArray();
        }

        /*****************************************************************/
    }
}
