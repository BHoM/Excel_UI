using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.oM.Geometry;
using BH.Engine.Reflection;

namespace BH.UI.Dragon
{
    public static partial class Create
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM object", Category = "Dragon")]
        public static object CreateGeometry(
                [ExcelArgument(Name = "geometry type")] string typeString,
                [ExcelArgument(Name = "property names (optional)")] object[] propNames,
                [ExcelArgument(Name = "property values (optional)")] object[] propValues)
        {

            if (propNames.Length != propValues.Length)
                return "Need to provide the same number of property names as property values";

            Type type = BH.Engine.Reflection.Create.Type(typeString);
            IGeometry geom = type.GetConstructor(Type.EmptyTypes).Invoke(new object[] { }) as IGeometry;

            string message;
            if (!InOutHelp.SetPropertyHelper(geom, propNames, propValues, out message))
                return message;

            Project.ActiveProject.Add(geom);
            return Project.ActiveProject.Add(geom).ToString();
        }
    }
}
