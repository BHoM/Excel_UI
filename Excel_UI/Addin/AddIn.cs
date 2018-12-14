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

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        private FormulaDataAccessor m_da;
        private Dictionary<string, CallerFormula> m_formulea;

        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/
        public void AutoOpen()
        {
            RegisterExcelMethods();
            RegisterBHoMMethods();
            
            //Hide error box showing methods not working properly
            if(!DebugConfig.ShowExcelDNALog)
                ExcelDna.Logging.LogDisplay.Hide();

            var app = ExcelDnaUtil.Application as Application;
            app.WorkbookBeforeSave += App_WorkbookBeforeSave;
            app.WorkbookOpen += App_WorkbookOpen;
        }

        private void App_WorkbookOpen(Workbook Wb)
        {
            List<string> json = new List<string>();
            try
            {
                _Worksheet newsheet = Wb.Sheets["BHoM_Data"];
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

        private void App_WorkbookBeforeSave(Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Project p = Project.ForWorkbook(Wb);

            _Worksheet newsheet;
            try
            {
                newsheet = Wb.Sheets["BHoM_Data"];
                try
                {
                    foreach (Range cell in newsheet.UsedRange)
                    {
                        cell.Value = "";
                    }
                }
                catch { }
            } catch
            {
                if (p.Empty) return;
                newsheet = Wb.Sheets.Add();
                newsheet.Name = "BHoM_Data";
            }
            if (p.Empty) return;

            newsheet.Visible = XlSheetVisibility.xlSheetHidden;
            int row = 1;
            var json = Project.ForWorkbook(Wb).Serialize();
            foreach (var obj in json)
            {
                Range cell = newsheet.Cells[row, 1];
                int c = 0;
                while (c < obj.Length)
                {
                    cell.Value = obj.Substring(c);
                    c += (cell.Value as string).Length;
                    cell = cell.Next;
                }
                row++;
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
                formula.Caller.Run();
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
