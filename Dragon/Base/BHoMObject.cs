using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using BHG = BHoM.Geometry;
using BHB = BHoM.Base;
using System.Reflection;

namespace Dragon.Base
{
    public static class BHoMObject
    {
        static BHoMObject()
        {
            Assembly basm = Assembly.GetAssembly(typeof(BHB.BHoMObject));
            string basmName = basm.FullName;
            string bhomName = typeof(BHB.BHoMObject).AssemblyQualifiedName;

            AssemblyNames = new Dictionary<string, string>();

            foreach (Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    string name = asm.GetName().Name;
                    if (name == "BHoM" || name.EndsWith("_oM"))
                    {
                        foreach (Type type in asm.GetTypes())
                        {
                            AssemblyNames[type.Name] = type.AssemblyQualifiedName;
                            AssemblyNames[type.FullName] = type.AssemblyQualifiedName;
                        }
                    }
                }
                catch(Exception e)
                {
                    Console.WriteLine(e);
                }
                
            }

            Console.WriteLine("Done");
        }

        private static Dictionary<string, string> AssemblyNames;

        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM point", Category = "Dragon")]
        public static object CreateObject(string typeString, object[] propNames, object[] propValues)
        {
            if (AssemblyNames.ContainsKey(typeString))
                typeString = AssemblyNames[typeString];

            Type type = Type.GetType(typeString);
            BHB.BHoMObject obj = BHB.BHoMObject.CreateInstance(type);

            try
            {
                int nb = Math.Min(propNames.Length, propValues.Length);
                for (int i = 0; i < nb; i++)
                {
                    BHB.BHoMJSON.ReadProperty(obj, (string)propNames[i], (string)propValues[i], BHoM.Global.Project.ActiveProject);
                }
            }
            catch
            {

            }

            return obj.ToJSON();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get the property of an object", Category = "Dragon")]
        public static object GetProperty(string BHoMObject, string property)
        {
            BHB.BHoMObject obj = BHB.BHoMObject.FromJSON(BHoMObject);
            System.Reflection.PropertyInfo prop = obj.GetType().GetProperty(property);
            return prop.GetValue(obj);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Set the property of an object", Category = "Dragon")]
        public static object SetProperty(string BHoMObject, object[] propNames, object[] propValues)
        {
            BHB.BHoMObject obj = BHB.BHoMObject.FromJSON(BHoMObject);
            int nb = Math.Min(propNames.Length, propValues.Length);
            for (int i = 0; i < nb; i++)
                BHB.BHoMJSON.ReadProperty(obj, (string)propNames[i], (string)propValues[i], BHoM.Global.Project.ActiveProject);

            return obj.ToJSON();
        }
    }
}
