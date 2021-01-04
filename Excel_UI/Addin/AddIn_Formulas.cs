/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2021, the respective contributors. All rights reserved.
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
using Microsoft.Office.Interop.Excel;
using BH.oM.UI;
using BH.UI.Base;
using BH.UI.Excel.Templates;
using BH.oM.Base;
using BH.UI.Excel.Components;

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
                // Register the caller formula with Excel if not already done
                if (m_Registered.Contains(caller.Function))
                {
                    if (callback != null)
                        ExcelAsyncUtil.QueueAsMacro(() => callback());
                }
                else
                {
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
                                ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
                            }

                            if (callback != null)
                                ExcelAsyncUtil.QueueAsMacro(() => callback());
                        }
                    });
                }

                // Save the caller info to the hidden sheet if needed
                if (saveToHiddenSheet)
                    SaveCallerToHiddenSheet(caller);
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
            for (int i = 2; i < 10000; i++)
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

                // Register the choices as objects if formula is a dropdown
                CallerValueListFormula valueList = formula as CallerValueListFormula;
                if (valueList != null)
                {
                    foreach (object choice in valueList.MultiChoiceCaller.Choices)
                    {
                        if (choice is IObject)
                            IAddObject(choice);
                    }
                }
            }
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static Tuple<Delegate, ExcelFunctionAttribute, List<object>> GetExcelDelegate(CallerFormula caller)
        {
            int nbInputs = caller.Caller.InputParams.Count;
            List<ParamInfo> inputs = caller.Caller.InputParams.ToList();
            ParameterExpression[] lambdaParams = inputs.Select(p => Expression.Parameter(typeof(object))).ToArray();
            NewArrayExpression array = Expression.NewArrayInit(typeof(object), lambdaParams);

            // Define the method to call depending on the number of outputs
            MethodCallExpression methodCall;
            if (caller.Caller.OutputParams.Count > 1)
            {
                List<ParamInfo> extraInputs = new List<ParamInfo>
                {
                    new ParamInfo { DataType = typeof(bool), DefaultValue = false, HasDefaultValue = true, IsRequired = false, Name = "_includeOutputNames", Description = "Include the name of the outputs" },
                    new ParamInfo { DataType = typeof(bool), DefaultValue = false, HasDefaultValue = true, IsRequired = false, Name = "_transposeOutputs", Description = "Transpose the resulting table (i.e. one output per row instead of per column)" }
                };

                inputs.AddRange(extraInputs);
                lambdaParams = lambdaParams.Concat(new ParameterExpression[] { Expression.Parameter(typeof(bool)), Expression.Parameter(typeof(bool)) }).ToArray();
                MethodInfo method = caller.GetType().GetMethod("RunWithOutputConfig");
                methodCall = Expression.Call(Expression.Constant(caller), method, array, lambdaParams[nbInputs], lambdaParams[nbInputs + 1]);
            }
            else
            {
                MethodInfo method = caller.GetType().GetMethod("Run");
                methodCall = Expression.Call(Expression.Constant(caller), method, array);
            }
            
            LambdaExpression lambda = Expression.Lambda(methodCall, lambdaParams);
            return new Tuple<Delegate, ExcelFunctionAttribute, List<object>>(
                lambda.Compile(),
                GetFunctionAttribute(caller),
                GetArgumentAttributes(inputs).ToList<object>()
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
                int nbFullName = argAttrs.Count;
                string argstring = argAttrs.Select(item => item.Name).Aggregate((a, b) => $"{a}, {b}");
                while (argstring.Length >= 254)
                {
                    nbFullName--;
                    ExcelArgumentAttribute arg = argAttrs[nbFullName];
                    bool isOptional = arg.Name.StartsWith("[");

                    arg.Description = "Full name: " + arg.Name + ". " + arg.Description;
                    arg.Name = string.Concat(arg.Name.Where(x => char.IsUpper(x)));
                    if (isOptional)
                        arg.Name = "[" + arg.Name + "]";

                    argstring = argAttrs.Select(item => item.Name).Aggregate((a, b) => $"{a}, {b}");
                }
            }

            return argAttrs;
        }

        /*******************************************/

        private static void SaveCallerToHiddenSheet(CallerFormula caller)
        {
            // Get the hidden worksheet
            Worksheet sheet = Sheet("BHoM_CallersHidden", true, true);
            if (sheet == null)
                return;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                lock(m_Mutex)
                {
                    // Get teh sheet ID
                    string sheetId = sheet.Cells[1, 1].Value as string;
                    if (sheetId == null)
                    {
                        sheetId = ToString(Guid.NewGuid());
                        sheet.Cells[1, 1].Value = sheetId;
                    }

                    // Make sure there is a list of saved callers for this worksheet
                    if (!m_SavedOnWorkbook.ContainsKey(sheetId))
                        m_SavedOnWorkbook[sheetId] = new HashSet<string>();

                    // Save the caller if not already done
                    if (!m_SavedOnWorkbook[sheetId].Contains(caller.Function))
                    {
                        m_SavedOnWorkbook[sheetId].Add(caller.Function);
                        int row = m_SavedOnWorkbook[sheetId].Count + 1;

                        sheet.Cells[row, 1].Value = caller.Caller.GetType().Name;
                        sheet.Cells[row, 2].Value = caller.Caller.Write();
                    } 
                } 
            });
            
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private static HashSet<string> m_Registered = new HashSet<string>();
        private static Dictionary<string, HashSet<string>> m_SavedOnWorkbook = new Dictionary<string, HashSet<string>>();
        private static object m_Mutex = new object();

        /*******************************************/
    }
}


