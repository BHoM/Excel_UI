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

using BH.Engine.Reflection;
using BH.Engine.Excel;
using BH.oM.UI;
using BH.UI.Base;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice.ExcelApi;
using System.Xml;
using BH.Engine.Serialiser;
using System.Reflection;
using System.Linq.Expressions;
using System.Collections;

namespace BH.UI.Excel.Templates
{
    public abstract partial class CallerFormula
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void Register()
        {
            Register(() => { });
        }

        /*******************************************/

        public void Register(System.Action callback)
        {
            lock (m_Mutex)
            {
                if (m_Registered.Contains(Function))
                {
                    callback();
                    return;
                }

                var formula = GetExcelDelegate();
                string function = Function;

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    lock (m_Mutex)
                    {
                        if (m_Registered.Contains(function))
                        {
                            ExcelAsyncUtil.QueueAsMacro(() => callback());
                            return;
                        }
                        ExcelIntegration.RegisterDelegates(
                            new List<Delegate>() { formula.Item1 },
                            new List<object> { formula.Item2 },
                            new List<List<object>> { formula.Item3 }
                        );
                        m_Registered.Add(function);
                        ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
                        ExcelAsyncUtil.QueueAsMacro(() => callback());
                    }
                });
            }
        }

        /*******************************************/

        public void EnqueueRegistration()
        {
            lock (m_Mutex)
            {
                if (m_Registered.Contains(Function))
                    return;

                var formula = GetExcelDelegate();
                m_RegistrationQueue.Enqueue(formula);
            }
        }

        /*******************************************/

        public static void RegisterQueue()
        {
            lock (m_Mutex)
            {
                if (m_RegistrationQueue.Count == 0)
                    return;
                var delegates = new List<Delegate>();
                var attrs = new List<object>();
                var paramAttrs = new List<List<object>>();
                while (m_RegistrationQueue.Count > 0)
                {
                    var current = m_RegistrationQueue.Dequeue();
                    if (m_Registered.Contains(current.Item2.Name))
                        continue;
                    delegates.Add(current.Item1);
                    attrs.Add(current.Item2);
                    paramAttrs.Add(current.Item3);
                }

                ExcelIntegration.RegisterDelegates(delegates, attrs, paramAttrs);
                foreach (ExcelFunctionAttribute attr in attrs)
                {
                    m_Registered.Add(attr.Name);
                }
            }
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected virtual Tuple<Delegate, ExcelFunctionAttribute, List<object>> GetExcelDelegate()
        {
            ParameterExpression[] lambdaParams = Caller.InputParams.Select(p => Expression.Parameter(typeof(object))).ToArray();
            NewArrayExpression array = Expression.NewArrayInit(typeof(object), lambdaParams);

            MethodInfo method = this.GetType().GetMethod("Run");
            MethodCallExpression methodCall = Expression.Call(Expression.Constant(this), method, array);
            LambdaExpression lambda = Expression.Lambda(methodCall, lambdaParams);

            return new Tuple<Delegate, ExcelFunctionAttribute, List<object>>(
                lambda.Compile(),
                GetFunctionAttribute(this),
                GetArgumentAttributes(Caller.InputParams).ToList<object>()
            );
        }

        /*******************************************/

        protected virtual ExcelFunctionAttribute GetFunctionAttribute(CallerFormula caller)
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

        protected virtual List<ExcelArgumentAttribute> GetArgumentAttributes(List<ParamInfo> rawParams)
        {
            List<ExcelArgumentAttribute> argAttrs = rawParams
                        .Select(p =>
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

                            return new ExcelArgumentAttribute()
                            {
                                Name = name,
                                Description = desc
                            };
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
        /**** Private Fields                    ****/
        /*******************************************/

        private static Queue<Tuple<Delegate, ExcelFunctionAttribute, List<object>>> m_RegistrationQueue = new Queue<Tuple<Delegate, ExcelFunctionAttribute, List<object>>>();
        private static HashSet<string> m_Registered = new HashSet<string>();
        private static object m_Mutex = new object();

        /*******************************************/
    }
}

