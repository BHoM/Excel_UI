using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;


namespace BH.UI.Dragon
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string sourceFolder = @"C:\Users\" + Environment.UserName + @"\AppData\Roaming\BHoM\BHoM_dlls";

           AssemblyName[] loadedAssemblies = ass.GetReferencedAssemblies();

            //Load all BHoM dlls on opening excel
            foreach (string path in Directory.GetFiles(sourceFolder, "*.dll"))
            {
                //Check that the dll is not allready loaded
                if(loadedAssemblies.Where(x => x.Name + ".dll" == Path.GetFileName(path)).Count() < 1)
                    Assembly.LoadFrom(path);
            }

        }

        public void AutoClose()
        {
        }
    }
}
