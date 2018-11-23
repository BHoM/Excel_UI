using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.Engine.Reflection;
using BH.oM.Base;
using BH.oM.Geometry;
using System.Linq.Expressions;
using BH.Adapter;
using BH.UI.Templates;
using BH.UI.Dragon.Templates;
using BH.UI.Dragon.Components;
using BH.UI.Dragon.Global;
using BH.UI.Global;
using BH.UI.Components;
using BH.Engine.Reflection.Convert;

namespace BH.UI.Dragon
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
            RegisterDragonMethods();
            RegisterBHoMMethods();
            
            //Hide error box showing methods not working properly
            if(!DebugConfig.ShowExcelDNALog)
                ExcelDna.Logging.LogDisplay.Hide();
        }

        /***************************************************/

        public void AutoClose()
        {
            
        }

        /*****************************************************************/
        /******* Private methods                            **************/
        /*****************************************************************/

        private void RegisterDragonMethods()
        {
            //Get out all the methods marked with the excel attributes
            IEnumerable<MethodInfo> allDragonMethods = ExcelIntegration.GetExportedAssemblies()
                .SelectMany(x => x.GetTypes().SelectMany(y => y.GetMethods(BindingFlags.Public | BindingFlags.Static)))
                .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() != null);

            List<MethodInfo> adapterMethods = new List<MethodInfo>();
            List<MethodInfo> otherMethods = new List<MethodInfo>();

            foreach (MethodInfo mi in allDragonMethods)
            {
                otherMethods.Add(mi);
            }
            List<object> fattrs = new List<object>();
            List<List<object>> attrs = new List<List<object>>();
            foreach (MethodInfo method in otherMethods)
            {
                var fa = method.GetCustomAttribute<ExcelFunctionAttribute>();
                fa.Name = "Dragon." + (fa.Name != null ? fa.Name : method.Name);

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
            m_da = new FormulaDataAccessor();
            var searcher = new FormulaSearchMenu(m_da);
            GlobalSearch.Attach(searcher);
            GlobalSearch.ItemSelected += GlobalSearch_ItemSelected;

            searcher.SetParent(null);
            Type fda = typeof(FormulaDataAccessor);
            Type callform = typeof(CallerFormula);
            Type[] constrtypes = new Type[] { fda };
            object[] args = new object[] { m_da };
            Type adapterType = typeof(BHoMAdapter);

            IEnumerable<MethodBase> methods = Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)).OrderBy(x => x.Name).SelectMany(x => x.GetConstructors());


            var adapterRegs = new List<Tuple<Delegate, ExcelFunctionAttribute, List<object>>>();
            foreach ( MethodBase adapter in methods)
            {
                var proxy = m_da.Wrap(adapter, () => GlobalSearch_ItemSelected(null, new oM.UI.ComponentRequest
                {
                    CallerType = typeof(CreateAdapterCaller),
                    SelectedItem = adapter
                }));
                adapterRegs.Add(proxy);
            }

            ExcelIntegration.RegisterDelegates(
                adapterRegs.Select(r => r.Item1).ToList(),
                adapterRegs.Select(r => r.Item2).ToList<object>(),
                adapterRegs.Select(r => r.Item3).ToList()
            );

            m_formulea = ExcelIntegration.GetExportedAssemblies()
                .SelectMany(a => a.GetTypes())
                .Where(t => t.Namespace == "BH.UI.Dragon.Components"
                            && callform.IsAssignableFrom(t))
                .Select(t => t.GetConstructor(constrtypes).Invoke(args) as CallerFormula)
                .ToDictionary(o => o.Caller.GetType().Name);
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
