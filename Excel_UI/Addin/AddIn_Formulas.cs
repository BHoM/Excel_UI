/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2020, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *
 *
 * The BHoM is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3.0 of the License, or
 * (at your option) any later version.
 *
 * The BHoM is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.
 */

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using System.Linq.Expressions;
using NetOffice.ExcelApi;
using BH.oM.UI;
using BH.UI.Base;
using BH.UI.Excel.Templates;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static void Register(CallerFormula caller, System.Action callback = null, bool saveToHiddenSheet = true)
        {
            lock (m_Mutex)
            {
                if (m_Registered.Contains(caller.Function))
                {
                    if (callback != null)
                        ExcelAsyncUtil.QueueAsMacro(() => callback());
                    return;
                }

                var formula = GetExcelDelegate(caller);
                string function = caller.Function;

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    lock (m_Mutex)
                    {
                        if (!m_Registered.Contains(function))
                        {
                            ExcelIntegration.RegisterDelegates(
                                new List<Delegate>() { formula.Item1 },
                                new List<object> { formula.Item2 },
                                new List<List<object>> { formula.Item3 }
                            );
                            m_Registered.Add(function);
                            if (saveToHiddenSheet)
                                SaveCallerToHiddenSheet(caller.Caller);
                            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
                        }

                        if (callback != null)
                            ExcelAsyncUtil.QueueAsMacro(() => callback());
                    }
                });
            }
        }

        /*******************************************/

        public static void RestoreFormulas()
        {
            // Get the hidden worksheet
            Worksheet sheet = Sheet("BHoM_CallersHidden", false);
            if (sheet == null)
            {
                Old_Restore(); // is it an old version of an Excel file ?
                return;
            }

            // Get all the formulas stored in teh BHoM_CallersHidden sheet
            for (int i = 1; i < 10000; i++)
            {
                // Recover the information about the formula
                string formulaName = sheet.Cells[i, 1].Value as string;
                string callerJson = sheet.Cells[i, 2].Value as string;
                if (formulaName == null || formulaName.Length == 0 || callerJson == null || callerJson.Length == 0)
                    break;

                // Register that formula from the json information
                CallerFormula formula = InstantiateCaller(formulaName);
                if (formula != null)
                {
                    formula.Caller.Read(callerJson);
                    Register(formula);
                }
            }

            // TODO: needs to wait for all formulas to be registered before returning
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static Tuple<Delegate, ExcelFunctionAttribute, List<object>> GetExcelDelegate(CallerFormula caller)
        {
            ParameterExpression[] lambdaParams = caller.Caller.InputParams.Select(p => Expression.Parameter(typeof(object))).ToArray();
            NewArrayExpression array = Expression.NewArrayInit(typeof(object), lambdaParams);

            MethodInfo method = caller.GetType().GetMethod("Run");
            MethodCallExpression methodCall = Expression.Call(Expression.Constant(caller), method, array);
            LambdaExpression lambda = Expression.Lambda(methodCall, lambdaParams);

            return new Tuple<Delegate, ExcelFunctionAttribute, List<object>>(
                lambda.Compile(),
                GetFunctionAttribute(caller),
                GetArgumentAttributes(caller.Caller.InputParams).ToList<object>()
            );
        }

        /*******************************************/

        private static ExcelFunctionAttribute GetFunctionAttribute(CallerFormula caller)
        {
            int limit = 254;
            string description = caller.Caller.Description;
            if (description.Length >= limit)
                description = description.Substring(0, limit - 1);
            return new ExcelFunctionAttribute()
            {
                Name = caller.Function,
                Description = description,
                Category = "BHoM." + caller.Caller.Category,
                IsMacroType = true
            };
        }

        /*******************************************/

        private static List<ExcelArgumentAttribute> GetArgumentAttributes(List<ParamInfo> rawParams)
        {
            List<ExcelArgumentAttribute> argAttrs = rawParams.Select(p =>
            {
                string name = p.HasDefaultValue ? $"[{p.Name}]" : p.Name;
                string postfix = string.Empty;
                if (p.HasDefaultValue)
                {
                    postfix += " [default: " +
                    (p.DefaultValue is string
                        ? $"\"{p.DefaultValue}\""
                        : p.DefaultValue == null
                            ? "null"
                            : p.DefaultValue.ToString()
                    ) + "]";
                }

                int limit = 253 - name.Length;
                string desc = p.Description + postfix;

                if (desc.Length >= limit)
                    desc = p.Description.Substring(limit - postfix.Length) + postfix;

                return new ExcelArgumentAttribute() { Name = name, Description = desc };
            }).ToList();

            if (argAttrs.Count() > 0)
            {
                string argstring = argAttrs.Select(item => item.Name).Aggregate((a, b) => $"{a}, {b}");
                if (argstring.Length >= 254)
                {
                    int i = 0;
                    argAttrs = argAttrs.Select(attr => new ExcelArgumentAttribute
                    {
                        Description = attr.Description,
                        Name = "arg" + i++
                    }).ToList();
                }
            }

            return argAttrs;
        }

        /*******************************************/

        private static void SaveCallerToHiddenSheet(Caller caller)
        {
            // Get the hidden worksheet
            Worksheet sheet = Sheet("BHoM_CallersHidden", true, true);
            if (sheet == null)
                return;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                lock(m_Mutex)
                {
                    int index = m_Registered.Count;
                    sheet.Cells[index, 1].Value = caller.GetType().Name;
                    sheet.Cells[index, 2].Value = caller.Write();
                } 
            });
            
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static HashSet<string> m_Registered = new HashSet<string>();
        private static object m_Mutex = new object();

        /*******************************************/
    }
}

