/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2024, the respective contributors. All rights reserved.
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

using BH.UI.Base;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Collections;
using BH.oM.UI;
using System.Reflection;

namespace BH.UI.Excel.Templates
{
    public abstract partial class CallerFormula
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public virtual Tuple<Delegate, ExcelFunctionAttribute, List<object>> GetExcelDelegate()
        {
            List<ParamInfo> inputs = Caller.InputParams.ToList();
            ParameterExpression[] lambdaParams = inputs.Select(p => Expression.Parameter(typeof(object))).ToArray();
            
            // Define the method to call depending on the number of outputs
            MethodCallExpression methodCall = GetMethodCall(ref inputs, ref lambdaParams);

            LambdaExpression lambda = Expression.Lambda(methodCall, lambdaParams);
            return new Tuple<Delegate, ExcelFunctionAttribute, List<object>>(
                lambda.Compile(),
                GetFunctionAttribute(),
                GetArgumentAttributes(inputs).ToList<object>()
            );
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected virtual MethodCallExpression GetMethodCall(ref List<ParamInfo> inputs, ref ParameterExpression[] lambdaParams)
        {
            int nbInputs = Caller.InputParams.Count;
            NewArrayExpression array = Expression.NewArrayInit(typeof(object), lambdaParams);

            MethodCallExpression methodCall;
            if (Caller.OutputParams.Count > 1)
            {
                List<ParamInfo> extraInputs = new List<ParamInfo>
                {
                    new ParamInfo { DataType = typeof(bool), DefaultValue = false, HasDefaultValue = true, IsRequired = false, Name = "_includeOutputNames", Description = "Include the name of the outputs" },
                    new ParamInfo { DataType = typeof(bool), DefaultValue = false, HasDefaultValue = true, IsRequired = false, Name = "_transposeOutputs", Description = "Transpose the resulting table (i.e. one output per row instead of per column)" }
                };

                inputs.AddRange(extraInputs);
                lambdaParams = lambdaParams.Concat(new ParameterExpression[] { Expression.Parameter(typeof(bool)), Expression.Parameter(typeof(bool)) }).ToArray();
                MethodInfo method = GetType().GetMethod("RunWithOutputConfig");
                methodCall = Expression.Call(Expression.Constant(this), method, array, lambdaParams[nbInputs], lambdaParams[nbInputs + 1]);
            }
            else
            {
                MethodInfo method = GetType().GetMethod("Run");
                methodCall = Expression.Call(Expression.Constant(this), method, array);
            }

            return methodCall;
        }

        /*******************************************/

        protected virtual ExcelFunctionAttribute GetFunctionAttribute()
        {
            int limit = 254;
            string description = Caller.Description;
            if (description.Length >= limit)
                description = description.Substring(0, limit - 1);
            return new ExcelFunctionAttribute()
            {
                Name = Function,
                Description = description,
                Category = "BHoM." + Caller.Category,
                IsMacroType = true
            };
        }

        /*******************************************/

        protected virtual List<ExcelArgumentAttribute> GetArgumentAttributes(List<ParamInfo> rawParams)
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
    }
}





