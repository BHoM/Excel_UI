using System;
using System.Collections;
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

            link.Push(toSend, key);

            return "ToMongo";
        }

        /*****************************************************************/
        [ExcelFunction(Description = "Test function", Category = "Mongo_Dragon")]
        public static string ToMongoDictionary(
            [ExcelArgument(Name = "objects")] object[,] objects,
            [ExcelArgument(Name = "key")] string key,
            [ExcelArgument(Name = "server address")] string server,
            [ExcelArgument(Name = "database name")] string database,
            [ExcelArgument(Name = "collection name")] string collection)
        {

            string json = "{";

            for (int i = 0; i < objects.GetLength(1); i++)
            {
                string header = objects[0, i].ToString();
                string valJson = "[";

                for (int j = 1; j < objects.GetLength(0); j++)
                {
                    valJson += ItemToMongo(objects[j, i]) + ",";
                }
                valJson = valJson.TrimEnd(',');
                valJson += "]";

                json += "\"" + header + "\":" + valJson + ",";
            }
            json = json.TrimEnd(',') + "}";

            MA.MongoLink link = new MA.MongoLink(server, database, collection);

            link.Push(new string[] { json }, key);

            return "ToMongo";
        }

        private static string ItemToMongo(object o)
        {
            Guid guid;
            BHB.BHoMObject bhO;
            if (Guid.TryParse(o.ToString(), out guid) && (bhO = BHG.Project.ActiveProject.GetObject(guid)) != null)
            {
                return BHB.JSONWriter.Write(bhO);
            }
            else
            {
                return BHB.JSONWriter.Write(o);
            }
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Test function", Category = "Mongo_Dragon")]
        public static object FromMongo(
            [ExcelArgument(Name = "query")] object[] query,
            [ExcelArgument(Name = "server address")] string server,
            [ExcelArgument(Name = "database name")] string database,
            [ExcelArgument(Name = "collection name")] string collection)
        {
            MA.MongoLink link = new MA.MongoLink(server, database, collection);
            IEnumerable<object> objects = link.Query(query.Select(x => x.ToString()).ToList());

            int count = objects.Count();

            //Check if object is dictionary
            if (objects.FirstOrDefault() is IDictionary && count == 1)
            {
                IDictionary dict = objects.First() as IDictionary;
                int colCount = dict.Keys.Count;
                int rowCount = 2;

                foreach (var key in dict.Keys)
                {
                    if (dict[key] is IList)
                        rowCount = Math.Max(rowCount, (dict[key] as IList).Count + 1);
                }

                var array = new object[rowCount, colCount];

                int n = 0;

                foreach (var key in dict.Keys)
                {
                    array[0, n] = key.ToString();
                    int m;
                    if (dict[key] is IList)
                    {
                        IList list = dict[key] as IList;
                        m = list.Count+1;
                        for (int i = 0; i < m-1; i++)
                        {
                            array[i + 1, n] = ItemFromMongo(list[i]);
                        }

                    }
                    else
                    {
                        array[1, n] = ItemFromMongo(dict[key]);
                        m = 2;
                    }
                    for (int i = m; i < rowCount; i++)
                    {
                        array[i, n] = "";
                    }

                    n++;
                }


                return XlCall.Excel(XlCall.xlUDF, "Resize", array);
            }
            else if (count == 1)
            {
                return ItemFromMongo(objects.First());
            }
            else
            {
                object[] array = objects.Select(x => ItemFromMongo(x)).ToArray();
                return XlCall.Excel(XlCall.xlUDF, "Resize", array);
            }


        }


        /*****************************************************************/


        private static object ItemFromMongo(object o)
        {
            double d;
            if (o == null)
            {
                return "Null Item";
            }
            if (o is BHB.BHoMObject)
            {
                BHB.BHoMObject bhO = o as BHB.BHoMObject;
                BHG.Project.ActiveProject.AddObject(bhO);
                return bhO.BHoM_Guid.ToString();
            }
            else if (double.TryParse(o.ToString(), out d))
            {
                return d;
            }
            else
            {
                return o.ToString();
            }
        }









        /*****************************************************************/

        [ExcelFunction(Description = "Test Array 2", Category = "Mongo_Dragon")]
        public static object TestArray2()
        {
            object[,] array = new object[,] { { 3.4, 8.9 }, { "Mongo", "rules" }, { "super duper", "Test" } };
            return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        }

        /****************************************************************/

       

    }
}
