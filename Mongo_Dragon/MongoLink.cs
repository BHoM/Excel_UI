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

            List<BHB.BHoMObject> toSend = new List<BHoM.Base.BHoMObject>();

            for (int i = 0; i < objects.Length; i++)
            {
                if (objects[i] is string)
                {
                     BHB.BHoMObject obj = BHG.Project.ActiveProject.GetObject(objects[i] as string);
                        if (obj != null)
                            toSend.Add(obj);
                }
            }

            
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
