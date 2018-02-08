using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.Engine.Reflection;

namespace BH.UI.Dragon
{
    public static class GenericMethodCall
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get the names of all the types of Methods", Category = "Dragon")]
        public static object GetAllMethods(
            [ExcelArgument(Name = "Optional type of method delcaration")] string type = null)
        {

            string[] methodNames;
            if (string.IsNullOrWhiteSpace(type))
                methodNames = Query.BHoMMethodList().Select(x => x.Name).ToArray();
            else
                methodNames = Query.BHoMMethodList().Where(x => x.DeclaringType.Name == type).Select(x => x.Name).ToArray();


            return XlCall.Excel(XlCall.xlUDF, "Resize", methodNames);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Calls a method with a specific name and set of arguments", Category = "Dragon")]
        public static object CallMethod(
                [ExcelArgument(Name = "MethodName")] string methodName,
               [ExcelArgument(Name = "Arguments")] object[] arguments )
        {

            List<MethodInfo> methods = Query.BHoMMethodList().Where(x => x.Name == methodName).ToList();

            if (methods.Count == 0)
                return "Method with the given name not found";

            arguments = arguments.Select(x => x.CheckAndGetObjectOrGeometry()).ToArray();

            object[] matchingArgs;
            MethodBase method;
            if (!MatchMethodAndAguments(methods, arguments, out matchingArgs, out method))
                return "Method matching the provided arguments not found";


            object result = null;
            try
            {
                result = method.Invoke(null, matchingArgs);
            }
            catch (Exception e)
            {
                return "Method call failed. Inner exception message: " + e.Message;
            }

            if (result is IEnumerable)
            {
                object[] arr = (result as IEnumerable<object>).Select(x => x.ReturnTypeHelper()).ToArray();
                return XlCall.Excel(XlCall.xlUDF, "Resize", arr);
            }
            else
                return result.ReturnTypeHelper();
            
        }

        /*****************************************************************/

        public static bool MatchMethodAndAguments(IEnumerable<MethodBase> methods, object[] arguments, out object[] matchingArgs, out MethodBase method)
        {
            bool found = false;
            matchingArgs = null;
            int argCount = arguments.Length;
            method = null;
            //Loop through all methods
            foreach (MethodBase info in methods)
            {
                ParameterInfo[] paramInfo = info.GetParameters();
                bool matching = true;
                matchingArgs = new object[paramInfo.Length];

                //Loop trhough all the arguments of the method
                for (int i = 0; i < paramInfo.Length; i++)
                {
                    if (i < argCount)
                    {
                        object match;
                        //Check of the parametertype matches the expected
                        if (CheckMatch(paramInfo[i], arguments[i], out match))
                            matchingArgs[i] = match;
                        else if ((arguments[i] == ExcelMissing.Value || arguments[i].GetType() == typeof(ExcelDna.Integration.ExcelEmpty)) && paramInfo[i].IsOptional) //Check for empty cells
                            matchingArgs[i] = paramInfo[i].RawDefaultValue;
                        else
                        {
                            matching = false;
                            break;
                        }
                    }
                    else
                    {
                        //Check if parameter is optional
                        if (paramInfo[i].IsOptional)
                        {
                            matchingArgs[i] = paramInfo[i].RawDefaultValue;
                        }
                        else
                        {
                            matching = false;
                            break;
                        }
                    }
                }

                if (matching)
                {
                    found = true;
                    method = info;
                    break;
                }

            }

            return found;
        }

        /*****************************************************************/

        private static bool CheckMatch(ParameterInfo pInfo, object obj, out object match)
        {
            Type pType = pInfo.ParameterType;

            //Check if same type
            if (pType == obj.GetType())
            {
                match = obj;
                return true;
            }

            if (obj is IExcelObject)
            {
                if (pType == (obj as IExcelObject).InnerObject.GetType())
                {
                    match = (obj as IExcelObject).InnerObject;
                    return true;
                }
            }

            //Check for string
            if (pType == typeof(string))
            {
                match = obj.ToString();
                return true;
            }

            //Check for int
            if (pType == typeof(int))
            {
                if (obj.IsNumeric())
                {
                    match = Convert.ToInt32(obj);
                    return true;
                }
                if (obj is string)
                {
                    int i;
                    if (int.TryParse(obj as string, out i))
                    {
                        match = i;
                        return true;
                    }
                }
            }

            //Check double
            if (pType == typeof(double))
            {
                if (obj.IsNumeric())
                {
                    match = (double)obj;
                    return true;
                }
                if(obj is string)
                {
                    double d;
                    if (double.TryParse(obj as string, out d))
                    {
                        match = d;
                        return true;
                    }
                }
            }

            //Check enum
            if (pType.IsEnum && obj is string)
            {
                match = Enum.Parse(pType, obj as string);
                return match != null;
            }

            match = null;
            return false;
        }

        /*****************************************************************/
    }
}
