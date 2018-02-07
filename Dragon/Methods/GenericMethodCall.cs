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
            MethodInfo method;
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

        private static bool MatchMethodAndAguments(List<MethodInfo> methods, object[] arguments, out object[] matchingArgs, out MethodInfo method)
        {
            bool found = false;
            matchingArgs = null;
            int argCount = arguments.Length;
            method = null;
            //Loop through all methods
            foreach (MethodInfo info in methods)
            {
                ParameterInfo[] paramInfo = info.GetParameters();
                bool matching = true;
                matchingArgs = new object[paramInfo.Length];

                //Loop trhough all the arguments of the method
                for (int i = 0; i < paramInfo.Length; i++)
                {
                    if (i < argCount)
                    {
                        //Check of the parametertype matches the expected
                        if (paramInfo[i].ParameterType == arguments[i].GetType())
                            matchingArgs[i] = arguments[i];
                        else if (arguments[i] == ExcelMissing.Value && paramInfo[i].IsOptional) //Check for empty cells
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
    }
}
