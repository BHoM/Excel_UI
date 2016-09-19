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
        /*****************************************************************/
        /**** Initial setup                                           ****/
        /*****************************************************************/

        static BHoMObject()
        {
            AssemblyNames = new Dictionary<string, string>();
            foreach (Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
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
        }
        private static Dictionary<string, string> AssemblyNames;


        /*****************************************************************/
        /**** Public methods                                          ****/
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM point", Category = "Dragon")]
        public static object CreateObject(
            [ExcelArgument(Name = "object type")] string typeString,
            [ExcelArgument(Name = "property names")] object[] propNames,
            [ExcelArgument(Name = "property values")] object[] propValues)
        {
            if (AssemblyNames.ContainsKey(typeString))
                typeString = AssemblyNames[typeString];

            Type type = Type.GetType(typeString);
            BHB.BHoMObject obj = BHB.BHoMObject.CreateInstance(type);

            try
            {
                int nb = Math.Min(propNames.Length, propValues.Length);
                for (int i = 0; i < nb; i++)
                    BHB.BHoMJSON.ReadProperty(obj, (string)propNames[i], (string)propValues[i], BHoM.Global.Project.ActiveProject);
            }
            catch {}

            return obj.ToJSON();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get the property of an object", Category = "Dragon")]
        public static object GetProperty(
            [ExcelArgument(Name = "object")] string BHoMObject,
            [ExcelArgument(Name = "property name")] string property)
        {
            BHB.BHoMObject obj = BHB.BHoMObject.FromJSON(BHoMObject);
            System.Reflection.PropertyInfo prop = obj.GetType().GetProperty(property);
            return prop.GetValue(obj);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Set the property of an object", Category = "Dragon")]
        public static object SetProperty(
            [ExcelArgument(Name = "object")] string BHoMObject,
            [ExcelArgument(Name = "property name")] object[] propNames,
            [ExcelArgument(Name = "property value")] object[] propValues)
        {
            BHB.BHoMObject obj = BHB.BHoMObject.FromJSON(BHoMObject);
            int nb = Math.Min(propNames.Length, propValues.Length);
            for (int i = 0; i < nb; i++)
                BHB.BHoMJSON.ReadProperty(obj, (string)propNames[i], (string)propValues[i], BHoM.Global.Project.ActiveProject);

            return obj.ToJSON();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Test Array", Category = "Dragon")]
        public static object TestArray()
        {
            object[,] array = new object[,] { { 3.4, 8.9 }, { "BHoM", "rules" } };
            return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        }
    }
}
