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
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using BH.Engine.Serialiser;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        private FormulaDataAccessor m_da;
        private Dictionary<string, CallerFormula> m_formulea;
        private CommandBarButton m_internalise;

        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        public void AutoOpen()
        {
            RegisterExcelMethods();
            RegisterBHoMMethods();
            AddInternalise();
            
            //Hide error box showing methods not working properly
            if(!DebugConfig.ShowExcelDNALog)
                ExcelDna.Logging.LogDisplay.Hide();

            var app = ExcelDnaUtil.Application as Application;
            app.WorkbookOpen += App_WorkbookOpen;
            ExcelDna.IntelliSense.IntelliSenseServer.Register();
        }

        private void AddInternalise()
        {
            var app = ExcelDnaUtil.Application as Application;
            var cmb = app.CommandBars["Cell"];
            var btn = cmb.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            btn.Tag = "Internalise_Data";
            btn.Caption = "Internalise Data";
            btn.Click += Internalise_Click;
            m_internalise = btn;
        }

        private void Internalise_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var app = ExcelDnaUtil.Application as Application;
            Range selected = app.Selection;

            foreach (Range objcell in selected)
            {
                string value;
                try
                {
                    value = (string)objcell.Value;
                    if (value == null || value.Length == 0) continue;
                } catch { continue; }

                Project proj = Project.ForIDs(new string[] { value });

                if (proj.Count((o) => !(o is Adapter.BHoMAdapter)) == 0) continue;
                proj.SaveData(app.ActiveWorkbook);

                objcell.Value = value;
            }
        }


        /*****************************************************************/

        private void App_WorkbookOpen(Workbook Wb)
        {
            List<string> json = new List<string>();
            _Worksheet newsheet;
            try
            {
                try
                {
                    newsheet = Wb.Sheets["BHoM_DataHidden"];
                } catch
                {
                    // Backwards compatibility
                    newsheet = Wb.Sheets["BHoM_Data"];
                }
                foreach (Range row in newsheet.UsedRange.Rows)
                {
                    string str = "";
                    try
                    {
                        Range cell = row.Cells[1, 1];
                        while (cell.Value != null && cell.Value is string && (cell.Value as string).Length > 0)
                        {
                            str += cell.Value;
                            cell = cell.Next;
                        }
                    }
                    catch { }
                    if (str.Length > 0)
                    {
                        json.Add(str);
                    }
                }
                Project.ActiveProject.Deserialize(json);
            }
            catch
            {
            }
        }

        /***************************************************/

        public void AutoClose()
        {
            
        }

        /*****************************************************************/
        /******* Private methods                            **************/
        /*****************************************************************/

        private void RegisterExcelMethods()
        {
            //Get out all the methods marked with the excel attributes
            IEnumerable<MethodInfo> allExcelMethods = ExcelIntegration.GetExportedAssemblies()
                .SelectMany(x => x.GetTypes().SelectMany(y => y.GetMethods(BindingFlags.Public | BindingFlags.Static)))
                .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() != null);

            List<MethodInfo> otherMethods = new List<MethodInfo>();

            foreach (MethodInfo mi in allExcelMethods)
            {
                otherMethods.Add(mi);
            }

            List<object> fattrs = new List<object>();
            List<List<object>> attrs = new List<List<object>>();
            foreach (MethodInfo method in otherMethods)
            {
                var fa = method.GetCustomAttribute<ExcelFunctionAttribute>();
                fa.Name = "BHoM." + (fa.Name != null ? fa.Name : method.Name);

                fattrs.Add(fa);
                attrs.Add(
                    method.GetParameters()
                        .Select(p => p.GetCustomAttribute<ExcelArgumentAttribute>() as object)
                        .ToList()
                );
            }
            ExcelIntegration.RegisterMethods(otherMethods,fattrs,attrs);
        }

        /*****************************************************************/
        private void RegisterBHoMMethods()
        {
            try
            {
                Compute.LoadAllAssemblies(Environment.GetEnvironmentVariable("APPDATA") + @"\BHoM\Assemblies");
                m_da = new FormulaDataAccessor();

                Type fda = typeof(FormulaDataAccessor);
                Type callform = typeof(CallerFormula);
                Type[] constrtypes = new Type[] { fda };
                object[] args = new object[] { m_da };
                m_formulea = ExcelIntegration.GetExportedAssemblies()
                    .SelectMany(a => a.GetTypes())
                    .Where(t => t.Namespace == "BH.UI.Excel.Components"
                                && callform.IsAssignableFrom(t))
                    .Select(t => t.GetConstructor(constrtypes).Invoke(args) as CallerFormula)
                    .ToDictionary(o => o.Caller.GetType().Name);

                var searcher = new FormulaSearchMenu(m_da, m_formulea);
                searcher.SetParent(null);

                searcher.ItemSelected += GlobalSearch_ItemSelected;

            }
            catch (Exception e)
            {
                Compute.RecordError(e.Message);
            }
        }

        private void GlobalSearch_ItemSelected(object sender, oM.UI.ComponentRequest e)
        {
            if (m_formulea.ContainsKey(e.CallerType.Name))
            {
                CallerFormula formula = m_formulea[e.CallerType.Name];
                formula.Caller.SetItem(e.SelectedItem);
                formula.Run();
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
    }
}
