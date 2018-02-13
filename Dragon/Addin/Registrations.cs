using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.Engine.Reflection;
using ExcelDna.Registration;
using BH.oM.Base;
using BH.oM.Geometry;
using System.Linq.Expressions;
using BH.Adapter;


namespace BH.UI.Dragon
{
    public static class Registration
    {

        /*****************************************************************/

        /// <summary>
        /// Registers the given functions with Excel-DNA.
        /// </summary>
        /// <param name="registrationEntries"></param>
        public static void RegisterBHFunctions(this IEnumerable<ExcelFunctionRegistration> registrationEntries)
        {
            var delList = new List<Delegate>();
            var attList = new List<object>();
            var argAttList = new List<List<object>>();
            foreach (var entry in registrationEntries)
            {
                try
                {
                    var del = entry.FunctionLambda.Compile();
                    var att = entry.FunctionAttribute;
                    var argAtt = new List<object>(entry.ParameterRegistrations.Select(pr => pr.ArgumentAttribute));

                    delList.Add(del);
                    attList.Add(att);
                    argAttList.Add(argAtt);
                }
                catch (Exception ex)
                {
                    //Logging.LogDisplay.WriteLine("Exception while registering method {0} - {1}", entry.FunctionAttribute.Name, ex.ToString());
                }
            }

            ExcelIntegration.RegisterDelegates(delList, attList, argAttList);
        }

        /*****************************************************************/

        //Creates a function registratiosn for excel from a list of methods
        public static List<ExcelFunctionRegistration> Registrations(this IEnumerable<MethodBase> methods, string prefix = "BH.")
        {

            List<ExcelFunctionRegistration> regs = new List<ExcelFunctionRegistration>();
            foreach (var group in methods.GroupBy(x => GetMethodName(x as dynamic)))
            {
                if (group.Count() == 1)
                {
                    regs.Add(ExcelFunctionRegistration(group.First(), prefix + GetMethodName(group.First() as dynamic)));
                }
                else
                {
                    foreach (MethodBase method in group)
                    {
                        string paramNames = ParamName(method);

                        regs.Add(ExcelFunctionRegistration(method, prefix + GetMethodName(method as dynamic) + paramNames));
                    }
                }
            }
            return regs;
        }

        /*****************************************************************/

        private static string GetMethodName(MethodInfo info)
        {
            return info.Name;
        }

        /*****************************************************************/
        private static string GetMethodName(ConstructorInfo info)
        {
            return info.DeclaringType.Name;
        }


        /*****************************************************************/

        private static string ParamName(MethodBase method)
        {
            string paramNames = "";
            foreach (ParameterInfo info in method.GetParameters())
            {
                if (typeof(IList).IsAssignableFrom(info.ParameterType))
                    paramNames += "_List" + info.ParameterType.GenericTypeArguments[0].Name;
                else if (typeof(IEnumerable).IsAssignableFrom(info.ParameterType) && info.ParameterType.IsGenericType)
                    paramNames += "_IEnum" + info.ParameterType.GenericTypeArguments[0].Name;
                else if (typeof(BHoMGroup<>).Name == info.ParameterType.Name && info.ParameterType.IsGenericType)
                    paramNames += "_Group" + info.ParameterType.GenericTypeArguments[0].Name;
                else
                    paramNames += "_" + info.ParameterType.Name;
            }

            return paramNames;
        }

        /*****************************************************************/
        private static ExcelFunctionRegistration ExcelFunctionRegistration(MethodBase method, string name)
        {
            var paramExprs = method.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();

            LambdaExpression lambda = GetLambdaExpression(method as dynamic, paramExprs, name);

            var allMethodAttributes = method.GetCustomAttributes(true);

            ExcelFunctionAttribute functionAttribute = null;

            foreach (var att in allMethodAttributes)
            {
                var funcAtt = att as ExcelFunctionAttribute;
                if (funcAtt != null)
                {
                    functionAttribute = funcAtt;
                    // At least ensure that name is set - from the method if need be.
                    if (string.IsNullOrEmpty(functionAttribute.Name))
                        functionAttribute.Name = name;
                }
            }
            // Check that ExcelFunctionAttribute has been set
            if (functionAttribute == null)
            {
                functionAttribute = new ExcelFunctionAttribute { Name = name };
            }

            IEnumerable<ExcelParameterRegistration> parameterRegistrations = method.GetParameters().Select(pi => new ExcelParameterRegistration(pi)).ToList();

            ExcelFunctionRegistration funcReg = new ExcelFunctionRegistration(lambda, functionAttribute, parameterRegistrations);
            funcReg.ReturnRegistration.CustomAttributes.AddRange(GetCustomReturnAttributes(method as dynamic));
            return funcReg;
        }

        /*****************************************************************/

        private static LambdaExpression GetLambdaExpression(MethodInfo info, List<ParameterExpression> paramExpres, string name)
        {
            return Expression.Lambda(Expression.Call(info, paramExpres), name, paramExpres);
        }

        /*****************************************************************/
        private static LambdaExpression GetLambdaExpression(ConstructorInfo info, List<ParameterExpression> paramExpres, string name)
        {
            NewExpression expression = Expression.New(info, paramExpres);
            return Expression.Lambda(expression, name, paramExpres);
        }

        /*****************************************************************/

        private static object[] GetCustomReturnAttributes(MethodInfo info)
        {
            return info.ReturnParameter.GetCustomAttributes(true);
        }

        /*****************************************************************/

        private static object[] GetCustomReturnAttributes(ConstructorInfo info)
        {
            //TODO: figure out what to return here
            return new object[] { };
        }

        /*****************************************************************/

    }
}
