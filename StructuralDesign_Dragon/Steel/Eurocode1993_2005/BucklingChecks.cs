using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using SSDT = StructuralDesign_Toolkit.Steel;

namespace StructuralDesign_Dragon.Steel.Eurocode1993_2005
{
    public static class BucklingChecks
    {
        //[ExcelFunction(Description = "Test function", Category = "Mongo_Dragon")]
        //public static string ToMongo(
        //    [ExcelArgument(Name = "objects")] object[] objects,
        //    [ExcelArgument(Name = "key")] string key,
        //    [ExcelArgument(Name = "server address")] string server,
        //    [ExcelArgument(Name = "database name")] string database,
        //    [ExcelArgument(Name = "collection name")] string collection)
        //{
        //    MA.MongoLink link = new MA.MongoLink(server, database, collection);

        //    List<BHB.BHoMObject> toSend = new List<BHoM.Base.BHoMObject>();

        //    for (int i = 0; i < objects.Length; i++)
        //    {
        //        if (objects[i] is string)
        //        {
        //            BHB.BHoMObject obj = BHG.Project.ActiveProject.GetObject(objects[i] as string);
        //            if (obj != null)
        //                toSend.Add(obj);
        //        }
        //    }

        //    link.Push(toSend, key);

        //    return "ToMongo";
        //}

        [ExcelFunction(Description = "Cm", Category = "StructuralDesign_Dragon")]
        public static double Cm(
            [ExcelArgument(Name = "Psi")] double psi,
            [ExcelArgument(Name = "Mh")] double mh,
            [ExcelArgument(Name = "Ms")] double ms)
        {

            return SSDT.Eurocode1993.BucklingChecks.Cm(psi, mh, ms);
        }
    }
}
