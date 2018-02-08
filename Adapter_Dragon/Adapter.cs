using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.Adapter;
using BH.Engine.Reflection;

namespace BH.UI.Dragon.Adapter
{
    public static class Adapter
    {

        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get available adapters", Category = "Dragon")]
        public static object GetAdapterTypes()
        {
            Type adapterType = typeof(BHoMAdapter);
            string[] adapterNames = Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)).Select(x => x.Name).ToArray();

            return XlCall.Excel(XlCall.xlUDF, "Resize", adapterNames);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get the parameters needed to create the specified adapter", Category = "Dragon")]
        public static object GetAdapterParameters(
            [ExcelArgument(Name = "adapter type type")] string typeString)
        {

            Type type = Query.AdapterTypeList().Where(x => x.Name == typeString).FirstOrDefault();

            if (type == null)
                return "No adapter of the specified type found. Please check the spelling";

            ConstructorInfo[] constrs = type.GetConstructors();

            if (constrs.Length < 1)
                return "No constructors found for the adapter";

            return XlCall.Excel(XlCall.xlUDF, "Resize", constrs[0].GetParameters().Select(x => x.Name).ToArray());
        }

        /*****************************************************************/
    }
}
