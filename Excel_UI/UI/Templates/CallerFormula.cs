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
    public abstract class CallerFormula
    {
        /*******************************************/
        /**** Events                            ****/
        /*******************************************/

        public event EventHandler OnRun;


        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public virtual string Category { get { return Caller.Category; } }

        public abstract string MenuRoot { get; }

        public virtual string Function
        {
            get
            {

                IEnumerable<ParamInfo> paramList = Caller.InputParams;
                bool hasParams = paramList.Count() > 0;
                string params_ = "";
                if (hasParams)
                {
                    params_ = "?by_" + paramList
                        .Select(p => p.DataType.ToText())
                        .Select(p => p.Replace("[]", "s"))
                        .Select(p => p.Replace("[,]", "Matrix"))
                        .Select(p => p.Replace("&", ""))
                        .Select(p => p.Replace("<", "Of"))
                        .Select(p => p.Replace(">", ""))
                        .Select(p => p.Replace(", ", "_"))
                        .Select(p => p.Replace("`", "_"))
                        .Aggregate((a, b) => $"{a}_{b}");
                }

                return GetName() + params_;
            }
        }

        public abstract Caller Caller { get; }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CallerFormula()
        {
            m_DataAccessor = new FormulaDataAccessor();
            Caller.SetDataAccessor(m_DataAccessor);
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public virtual string GetName()
        {
            if (Caller.SelectedItem != null && Caller.SelectedItem is MethodBase)
            {
                Type decltype = ((MethodBase)Caller.SelectedItem).DeclaringType;
                string ns = decltype.Namespace;
                if (ns.StartsWith("BH"))
                    ns = ns.Split('.').Skip(2).Aggregate((a, b) => $"{a}.{b}");
                return decltype.Name + "." + ns + "." + Caller.Name;
            }
            return Category + "." + Caller.Name;
        }

        /*******************************************/

        public void FillFormula(oM.Excel.Reference cell)
        {
            Register(() => Fill(cell));
        }

        /*******************************************/

        public virtual string GetRibbonXml()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement menu = doc.CreateElement("dynamicMenu");
            menu.SetAttribute("id", Caller.GetType().Name);
            menu.SetAttribute("getImage", "GetImage");
            menu.SetAttribute("label", MenuRoot);
            menu.SetAttribute("screentip", MenuRoot);
            menu.SetAttribute("supertip", Caller.Description);
            menu.SetAttribute("getContent", "GetContent");
            return menu.OuterXml;
        }

        /*******************************************/

        public virtual string GetInnerRibbonXml()
        {
            Caller.SelectedItem = null;
            m_Menu = SelectorMenuUtil.ISetExcelSelectorMenu(Caller.GetItemSelectorMenu());
            m_Menu.RootName = Caller.GetType().Name;
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("root");
            Caller.AddToMenu(root);
            XmlElement menu = root.FirstChild as XmlElement;
            if (menu == null)
                return "";
            menu.RemoveAllAttributes();
            menu.SetAttribute("xmlns", "http://schemas.microsoft.com/office/2006/01/customui");
            return root.InnerXml;
        }

        /*******************************************/

        public virtual void Select(string id)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                m_Menu.Select(id);
                FillFormula(Engine.Excel.Query.Selection());
            });
        }

        /*******************************************/

        public Tuple<Delegate, ExcelFunctionAttribute, List<object>> GetExcelDelegate()
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

        public virtual object Run(object[] inputs)
        {
            //Clear current events
            Engine.Reflection.Compute.ClearCurrentEvents();

            // Run the caller
            m_DataAccessor.SetInputs(inputs.ToList(), Caller.InputParams.Select(x => x.DefaultValue).ToList());
            Caller.Run();
            object result = m_DataAccessor.GetOutputs();

            // Handle possible errors
            var errors = Engine.Reflection.Query.CurrentEvents().Where(e => e.Type == oM.Reflection.Debugging.EventType.Error);
            if (errors.Count() > 0)
                Engine.Excel.Query.Caller().Note(errors.Select(e => e.Message).Aggregate((a, b) => a + "\n" + b));
            else
                Engine.Excel.Query.Caller().Note("");

            // Trigger event (why do we need this ?)
            OnRun?.Invoke(this, null);

            // Return result
            return result;
        }

        /*******************************************/

        public void Register()
        {
            Register(() => { });
        }

        /*******************************************/

        public void Register(System.Action callback)
        {
            lock(m_Mutex)
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
                        ExcelAsyncUtil.QueueAsMacro(()=>callback());
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
            lock(m_Mutex)
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

        protected virtual void Fill(oM.Excel.Reference cell)
        {
            System.Action callback = () => { };
            var cellcontents = "=" + Function;
            if (Caller.InputParams.Count == 0)
            {
                cellcontents += "()";
            }
            else
            {
                callback = () =>
                {
                    bool isNumlock = System.Windows.Forms.Control.IsKeyLocked(System.Windows.Forms.Keys.NumLock);
                    Application.GetActiveInstance().SendKeys("{F2}{(}", true);
                    if (isNumlock)
                        Application.GetActiveInstance().SendKeys("{NUMLOCK}", true);
                };
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var xlRef = cell.ToExcel();
                    XlCall.Excel(XlCall.xlcFormula, cellcontents, xlRef);
                    using (new ExcelSelectionHelper(xlRef))
                    {
                        callback();
                    }
                }
                catch { }
            });
        }

        /*******************************************/

        private ExcelFunctionAttribute GetFunctionAttribute(CallerFormula caller)
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

        private List<ExcelArgumentAttribute> GetArgumentAttributes(List<ParamInfo> rawParams)
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

        private IExcelSelectorMenu m_Menu;
        protected FormulaDataAccessor m_DataAccessor = null;

        private static Queue<Tuple<Delegate, ExcelFunctionAttribute, List<object>>> m_RegistrationQueue = new Queue<Tuple<Delegate, ExcelFunctionAttribute, List<object>>>();
        private static HashSet<string> m_Registered = new HashSet<string>();
        private static object m_Mutex = new object();

        /*******************************************/
    }
}

