/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2018, the respective contributors. All rights reserved.
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
using BH.Engine.Reflection;
using BH.oM.Base;
using System.Linq.Expressions;
using BH.UI.Templates;
using BH.UI.Excel.Templates;
using BH.UI.Excel.Components;
using BH.UI.Excel.Global;
using BH.UI.Global;
using BH.UI.Components;
using BH.Engine.Serialiser;
using System.Runtime.InteropServices;
using NetOffice.ExcelApi;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.ExcelApi.Enums;
using System.Drawing;
using System.Xml;

namespace BH.UI.Excel
{
    public class AddIn : IExcelAddIn
    {
        private Dictionary<string, CallerFormula> m_formulea;
        private List<CommandBar> m_menus;
        private Application m_application;
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        public void AutoOpen()
        {
            Instance = this;

            ExcelDna.IntelliSense.IntelliSenseServer.Install();

            m_application = Application.GetActiveInstance();
            using (Engine.Excel.Profiling.Timer timer = new Engine.Excel.Profiling.Timer("open"))
            {
                m_menus = new List<CommandBar>();
                m_menus.Add(m_application.CommandBars["Cell"]);

                Type callform = typeof(CallerFormula);

                Type[] constrtypes = new Type[] { };
                object[] args = new object[] { };

                m_formulea = ExcelIntegration.GetExportedAssemblies()
                    .SelectMany(a => a.GetTypes())
                    .Where(t => t.Namespace == "BH.UI.Excel.Components"
                                && callform.IsAssignableFrom(t))
                    .Select(t => t.GetConstructor(constrtypes).Invoke(args) as CallerFormula)
                    .ToDictionary(o => o.Caller.GetType().Name);

                m_application.WorkbookOpenEvent += App_WorkbookOpen;
            }
            
        }

        private void AddInternalise()
        {
            foreach (var cmb in m_menus)
            {
                var btn = cmb.Controls.Add(MsoControlType.msoControlButton,null,null,null, true) as CommandBarButton; 
                btn.Tag = "Internalise_Data";
                btn.Caption = "Internalise Data";
                btn.ClickEvent += Internalise_Click;
                btn.Dispose(false);
            }
        }

        private void Internalise_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Range selected = null;
            CancelDefault = true;

            try
            {
                    selected = m_application.Selection as Range;

                    foreach (Range objcell in selected)
                    {
                        string value;
                        try
                        {
                            value = (string)objcell.Value;
                            if (value == null || value.Length == 0) continue;
                        }
                        catch { continue; }

                        Project proj = Project.ForIDs(new string[] { value });

                        if (proj.Count((o) => !(o is Adapter.BHoMAdapter)) == 0) continue;
                        proj.SaveData(m_application.ActiveWorkbook);

                        objcell.Value = value;
                        objcell.Dispose();
                    }
            }
            finally
            {
                if (selected != null) selected.Dispose();
            }
        }


        /*****************************************************************/

        private void App_WorkbookOpen(Workbook Wb)
        {
            List<string> json = new List<string>();
            Sheets sheets = null;
            _Worksheet newsheet = null;
            Range used = null;
            Range cell = null;
            Range next = null;
            try
            {
                sheets = Wb.Sheets;

                if (sheets.OfType<Worksheet>()
                    .FirstOrDefault(s => s.Name == "BHoM_Used") != null)
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        InitBHoMAddin();
                        foreach (Worksheet sheet in sheets.OfType<Worksheet>())
                        {
                            try
                            {
                                bool before = sheet.EnableCalculation;
                                sheet.EnableCalculation = false;
                                sheet.Calculate();
                                sheet.EnableCalculation = before;
                            }
                            finally
                            {
                                sheet.Dispose();
                            }
                        }
                    });
                }

                try
                {
                    newsheet = sheets["BHoM_DataHidden"] as Worksheet;
                } catch
                {
                    // Backwards compatibility
                    newsheet = sheets["BHoM_Data"] as Worksheet;
                }
                used = newsheet.UsedRange;
                foreach (Range row in used.Rows)
                {
                    string str = "";
                    try
                    {
                        cell = row.Cells[1, 1];
                        while (cell.Value != null && cell.Value is string && (cell.Value as string).Length > 0)
                        {
                            str += cell.Value;
                            next = cell.Next;
                            cell.Dispose();
                            cell = next;
                        }
                    }
                    catch { }

                    if (str.Length > 0)
                    {
                        json.Add(str);
                    }

                    row.Dispose();
                }
                Project.ActiveProject.Deserialize(json);

            }
            finally
            {
                if (newsheet != null) newsheet.Dispose();
                if (cell != null) cell.Dispose();
                if (used != null) used.Dispose();
                if (next != null) next.Dispose();
            }
        }
        
        /***************************************************/

        public void AutoClose()
        {
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall();
        }

        /*****************************************************************/
        /******* Private methods                            **************/
        /*****************************************************************/


        private void RegisterBHoMMethods()
        {
            try
            {
                Compute.LoadAllAssemblies(Environment.GetEnvironmentVariable("APPDATA") + @"\BHoM\Assemblies");

                var searcher = new FormulaSearchMenu(m_formulea);
                searcher.SetParent(null);

                searcher.ItemSelected += Formula_ItemSelected;
                globalSearch.ItemSelected += GlobalSearch_ItemSelected;
            }
            catch (Exception e)
            {
                Compute.RecordError(e.Message);
            }
        }

        private void Formula_ItemSelected(object sender, oM.UI.ComponentRequest e)
        {
            if (m_formulea.ContainsKey(e.CallerType.Name))
            {
                CallerFormula formula = m_formulea[e.CallerType.Name];
                formula.Caller.SetItem(e.SelectedItem);
                formula.Run();
                FlagUsed();
            }
        }
        
        /*****************************************************************/
        private void FlagUsed()
        {
            Workbook Wb = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            try
            {
                Wb = m_application.ActiveWorkbook;
                sheets = Wb.Worksheets;
                if (sheets.OfType<Worksheet>()
                    .FirstOrDefault(s => s.Name == "BHoM_Used") == null)
                {
                    sheet = sheets.Add() as Worksheet;
                    sheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                    sheet.Name = "BHoM_Used";
                }
            } finally
            {
                if (Wb != null) Wb.Dispose();
                if (sheet != null) sheet.Dispose();
                if (sheets != null) sheets.Dispose();
            }
        }

        /*****************************************************************/
      
        [ExcelCommand(ShortCut = "^B")]
        public static void InitGlobalSearch()
        {
            var control = new System.Windows.Forms.ContainerControl();
            globalSearch.SetParent(control);
        }
        private static SearchMenu globalSearch = new SearchMenu_WinForm();
        
        public static void EnableBHoM(Action<bool> callback)
        {
            ExcelAsyncUtil.QueueAsMacro(() => callback(Instance.InitBHoMAddin()));
        }

        public bool InitBHoMAddin()
        {
            if (initialised) return true;
            RegisterBHoMMethods();
            ExcelDna.Registration.ExcelRegistration.RegisterCommands(ExcelDna.Registration.ExcelRegistration.GetExcelCommands());
            AddInternalise();
            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
            initialised = true;
            return true;
        }

        
        private bool initialised = false;
        public static bool Enabled { get { return Instance.initialised; } }

        public static CallerFormula GetCaller(string caller)
        {
            if (Instance.m_formulea.ContainsKey(caller))
            {
                return Instance.m_formulea[caller];
            }
            return null;
        }

        /*****************************************************************/

        private void GlobalSearch_ItemSelected(object sender, oM.UI.ComponentRequest e)
        {

            if (m_formulea.ContainsKey(e.CallerType.Name))
            {
                CallerFormula formula = m_formulea[e.CallerType.Name];
                formula.Caller.SetItem(e.SelectedItem);
                formula.FillFormula();
            }
        }


        /*****************************************************************/

        private static bool IsNullMissingOrEmpty(object obj)
        {
            if (obj == null)
                return true;

            if (obj == ExcelMissing.Value)
                return true;

            if (obj == ExcelEmpty.Value)
                return true;

            if (obj is string && string.IsNullOrWhiteSpace(obj as string))
                return true;

            return false;
        }

        public static string GetRibbonXml()
        {
            Dictionary<string, XmlElement> groups = new Dictionary<string, XmlElement>();
            Dictionary<string, Dictionary<int, XmlElement>> boxes = new Dictionary<string, Dictionary<int, XmlElement>>();
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("root");
            doc.AppendChild(root);
            foreach(CallerFormula caller in Instance.m_formulea.Values)
            {
                try
                {
                    XmlElement group;
                    groups.TryGetValue(caller.Category, out group);
                    if (group == null)
                    {
                        group = (XmlElement)root.AppendChild(doc.CreateElement("group"));
                        group.SetAttribute("id", caller.Category);
                        group.SetAttribute("label", caller.Category);
                        group.SetAttribute("getVisible", "GetVisible");
                        groups.Add(caller.Category, group);
                        boxes.Add(caller.Category, new Dictionary<int, XmlElement>());
                    }
                    if (!boxes[caller.Category].ContainsKey(caller.Caller.GroupIndex))
                        boxes[caller.Category].Add(caller.Caller.GroupIndex, doc.CreateElement("box"));

                    XmlElement box = boxes[caller.Category][caller.Caller.GroupIndex];
                    box.SetAttribute("id", caller.Category+"-group" + caller.Caller.GroupIndex);
                    box.SetAttribute("boxStyle", "vertical");

                    XmlDocument tmp = new XmlDocument();
                    tmp.LoadXml(caller.GetRibbonXml());
                    box.AppendChild(doc.ImportNode(tmp.DocumentElement, true));
                } catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            foreach(var kvp in boxes)
            {
                List<int> ordered = kvp.Value.Keys.ToList();
                ordered.Sort();
                foreach(int i in ordered)
                {
                    groups[kvp.Key].AppendChild(kvp.Value[i]);
                }
            }
            return root.InnerXml;
        }

        private static AddIn Instance { get; set; }
    }
}
