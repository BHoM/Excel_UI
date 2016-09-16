using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using BHG = BHoM.Geometry;
using MA = Mongo_Adapter;
using BHB = BHoM.Base;

namespace Dragon
{
    public static class Functions
    {
        /*****************************************************************/
        /****  Base                                                   ****/
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM point", Category = "Dragon")]
        public static object BHoMObject(string typeString, object[] propNames, object[] propValues)
        {
            Type type = Type.GetType(typeString);
            BHB.BHoMObject obj = BHB.BHoMObject.CreateInstance(type);

            int nb = Math.Min(propNames.Length, propValues.Length);
            for (int i = 0; i < nb; i++)
            {
                BHB.BHoMJSON.ReadProperty(obj, (string)propNames[i], (string)propValues[i], BHoM.Global.Project.ActiveProject);
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

        /*****************************************************************/
        /****  Geometry                                               ****/
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM point", Category = "Dragon")]
        public static object Point(double x, double y, double z)
        {
            return new BHG.Point(x, y, z).ToJSON();
        }

        /*****************************************************************/




    }
}
