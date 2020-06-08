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

using BH.Engine.Excel;
using BH.Engine.Reflection;
using BH.oM.Base;
using BH.oM.UI;
using BH.UI.Templates;
using ExcelDna.Integration;
using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace BH.UI.Excel.Templates
{
    public abstract class CallerFormula
    {
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
                if (Caller.SelectedItem is Type)
                {
                    params_ = "?by_Properties";
                }
                return GetFormulaName() + params_;
            }
        }

        public abstract Caller Caller { get; }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CallerFormula()
        {
            Caller.SetDataAccessor(new FormulaDataAccessor());
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public virtual string GetFormulaName()
        {
            Type declaringType = null;
            string nameSpace = "";
            if (Caller is MethodCaller && Caller.SelectedItem != null)
            {
                if (Caller.SelectedItem is Type)
                {
                    declaringType = (Caller as MethodCaller).OutputParams.First().DataType;
                    if (typeof(IObject).IsAssignableFrom(declaringType))
                    {
                        nameSpace = declaringType.Namespace;
                    }
                }
                else
                {
                    declaringType = (Caller as MethodCaller).Method.DeclaringType;
                }
                if (declaringType != null) nameSpace = declaringType.Namespace;
            }
            if (nameSpace.StartsWith("BH") && declaringType != null)
            {
                nameSpace = nameSpace.Split('.').Skip(2).Aggregate((a, b) => $"{a}.{b}");
                if (Caller.SelectedItem is Type)
                {
                    return "BH." + Category + ".Create." + nameSpace + "." + declaringType.Name;
                }
                else
                {                   
                    return "BH." + Category + "." + declaringType.Name + "." + nameSpace + "." + Caller.Name;
                }
            }
            else
            {
                return Caller.GetFullName();
            }
        }

        /*******************************************/

        public void FillFormula(oM.Excel.Reference cell)
        {
            Register(() => Fill(cell));
        }

        /*******************************************/

        public bool Run()
        {
            bool success = Excecute();
            OnRun?.Invoke(this, null);
            return success;
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
            m_Menu = SelectorMenuUtil.ISetExcelSelectorMenu(Caller.Selector);
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
            var accessor = Caller.DataAccessor as FormulaDataAccessor;
            object item = Caller.SelectedItem;
            return accessor.Wrap(this, () => RunItem(item));
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
        /**** Events                            ****/
        /*******************************************/

        public event EventHandler OnRun;

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        protected virtual bool Excecute()
        {
            return Caller.Run();
        }

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
                    Application.GetActiveInstance().SendKeys("{F2}{(}", true);
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

        private void RunItem(object item)
        {
            Caller.SetItem(item);
            Run();
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private IExcelSelectorMenu m_Menu;
        private static Queue<Tuple<Delegate, ExcelFunctionAttribute, List<object>>> m_RegistrationQueue =
            new Queue<Tuple<Delegate, ExcelFunctionAttribute, List<object>>>();
        private static HashSet<string> m_Registered = new HashSet<string>();
        private static object m_Mutex = new object();

        /*******************************************/
    }
}

