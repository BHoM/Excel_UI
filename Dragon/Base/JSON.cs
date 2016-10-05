using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BHB = BHoM.Base;
using BHG = BHoM.Global;

namespace Dragon.Base
{
    public static class JSON
    {
        /*****************************************************************/

        [ExcelFunction(Description = "Write package to send to Flux", Category = "Dragon")]
        public static string WritePackage(
            [ExcelArgument(Name = "objects")] object[] objectIds,
            [ExcelArgument(Name = "password (optional)")] string password = "")
        {
            Guid guid;
            List<BHB.BHoMObject> list = objectIds.Select(x =>  Guid.TryParse(x as string, out guid)? BHG.Project.ActiveProject.GetObject(guid): null).Where(x=> x!= null).ToList();
            return BHB.BHoMJSON.WritePackage(list, password);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Read package comming from Flux", Category = "Dragon")]
        public static object ReadPackage(
            [ExcelArgument(Name = "package")] string package,
            [ExcelArgument(Name = "password (optional)")] string password = "")
        {
            object[] objects = BHB.BHoMJSON.ReadPackage(package, password).Select(x => x.ToJSON()).ToArray();
            return XlCall.Excel(XlCall.xlUDF, "Resize", objects);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Convert the object to a JSON string", Category = "Dragon")]
        public static string ToJSON(
            [ExcelArgument(Name = "object")] string objectId)
        {
            BHB.BHoMObject obj = BHG.Project.ActiveProject.GetObject(objectId);
            if (obj == null)
                return "";
            else
                return obj.ToJSON();
        }
    }
}
