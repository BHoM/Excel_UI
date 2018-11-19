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
using BH.UI.Dragon.UI.Templates;

namespace BH.UI.Dragon
{
    public partial class AddIn : IExcelAddIn
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/
        public void AutoOpen()
        {
            LoadBHomAssemblies();
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
        private void LoadBHomAssemblies()
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string sourceFolder = @"C:\Users\" + Environment.UserName + @"\AppData\Roaming\BHoM\Assemblies";

            List<string> loadedAssemblies = ass.GetReferencedAssemblies().Select(x => x.Name + ".dll").ToList();

            //Load all BHoM dlls on opening excel
            foreach (string path in Directory.GetFiles(sourceFolder, "*.dll"))
            {
                try
                {
                    //Check that the dll is not allready loaded
                    if (loadedAssemblies.Where(x => x == Path.GetFileName(path)).Count() < 1)
                        Assembly.LoadFrom(path);
                }
                catch
                {
                    
                }
            }
        }

        /*****************************************************************/
        private void RegisterDragonMethods()
        {
            //Get out all the methods marked with the excel attributes
            IEnumerable<MethodInfo> allDragonMethods = ExcelIntegration.GetExportedAssemblies()
                .SelectMany(x => x.GetTypes().SelectMany(y => y.GetMethods(BindingFlags.Public | BindingFlags.Static)))
                .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() != null);

            List<MethodInfo> adapterMethods = new List<MethodInfo>();
            List<MethodInfo> otherMethods = new List<MethodInfo>();

            Type adapterType = typeof(Dragon.Adapter.Adapter);

            foreach (MethodInfo mi in allDragonMethods)
            {
                if (mi.DeclaringType == adapterType)
                    adapterMethods.Add(mi);
                else
                    otherMethods.Add(mi);
            }
        }

        /*****************************************************************/
        private void RegisterBHoMMethods()
        {
            IEnumerable<MethodBase> methods = Query.BHoMMethodList();
            var callers = new List<IFormulaCaller>();
            foreach (MethodBase method in methods)
            {
                try
                {
                    callers.Add(new MethodFormulaCaller(method));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            }
            try
            {
                ExcelIntegration.RegisterDelegates(
                    callers.Select(caller => caller.ExcelMethod).ToList(),
                    callers.Select(caller => caller.FunctionAttribute).Cast<object>().ToList(),
                    callers.Select(caller => caller.ExcelParams.Cast<object>().ToList()).ToList()
                );
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
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
