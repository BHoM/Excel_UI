using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using MA = Mongo_Adapter;
using BHB = BHoM.Base;
using BHG = BHoM.Global;


namespace Mongo_Dragon
{
    public static class MongoLink
    {
        /*****************************************************************/

        [ExcelFunction(Description = "Test function", Category = "Mongo_Dragon")]
        public static string ToMongo(
            [ExcelArgument(Name = "objects")] object[] objects,
            [ExcelArgument(Name = "key")] string key,
            [ExcelArgument(Name = "server address")] string server,
            [ExcelArgument(Name = "database name")] string database,
            [ExcelArgument(Name = "collection name")] string collection)
        {
            MA.MongoLink link = new MA.MongoLink(server, database, collection);

            List<BHB.BHoMObject> toSend = objects.Select(x => BHG.Project.ActiveProject.GetObject(x as string)).ToList();
            link.SaveObjects(toSend, key);

            return "ToMongo";
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Test function", Category = "Mongo_Dragon")]
        public static object FromMongo(
            [ExcelArgument(Name = "query")] string query,
            [ExcelArgument(Name = "server address")] string server,
            [ExcelArgument(Name = "database name")] string database,
            [ExcelArgument(Name = "collection name")] string collection)
        {
            MA.MongoLink link = new MA.MongoLink(server, database, collection);
            IEnumerable<BHB.BHoMObject> objects = link.GetObjects(query);

            object[] array = objects.Select(x => x.BHoM_Guid.ToString()).ToArray();
            return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Test Array 2", Category = "Mongo_Dragon")]
        public static object TestArray2()
        {
            object[,] array = new object[,] { { 3.4, 8.9 }, { "Mongo", "rules" } };
            return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        }
    }
}
