using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.Adapter;
using BH.Engine.Reflection;
using BH.oM.Base;
using BH.oM.Queries;

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
            [ExcelArgument(Name = "Adapter type")] string typeString)
        {

            Type type = Query.AdapterTypeList().Where(x => x.Name == typeString).FirstOrDefault();

            if (type == null)
                return "No adapter of the specified type found. Please check the spelling";

            ConstructorInfo[] constrs = type.GetConstructors();

            if (constrs.Length < 1)
                return "No constructors found for the adapter";

            return XlCall.Excel(XlCall.xlUDF, "Resize", constrs.OrderByDescending(x => x.GetParameters().Length).First().GetParameters().Select(x => x.Name).ToArray());
        }

        /*****************************************************************/
        [ExcelFunction(Description = "Create an adapter", Category = "Dragon")]
        public static object CreateAdapter(
            [ExcelArgument(Name = "Adapter type")] string typeString,
            [ExcelArgument(Name = "ConstructionParameters")] object[] arguments
            )
        {
            Type type = Query.AdapterTypeList().Where(x => x.Name == typeString).FirstOrDefault();

            if (type == null)
                return "No adapter of the specified type found. Please check the spelling";


            BHoMAdapter adapter = null;

            object[] matchingArgs;
            MethodBase method;
            if (!GenericMethodCall.MatchMethodAndAguments(type.GetConstructors().OrderByDescending(x => x.GetParameters().Length), arguments, out matchingArgs, out method))
                return "Method matching the provided arguments not found";

            ConstructorInfo constr = method as ConstructorInfo;

            try
            {
                adapter = constr.Invoke(matchingArgs) as BHoMAdapter;
            }
            catch (Exception e)
            {
                return "Failed creating adapter. Please check your arguments. Error message: " + e.Message;
            }

            return Project.ActiveProject.AddAdapter(adapter).ToString();
        }

        /*****************************************************************/
        [ExcelFunction(Description = "Create an adapter", Category = "Dragon")]
        public static object Push(
            [ExcelArgument(Name = "Adapter")] string adapterId,
            [ExcelArgument(Name = "Objects to push")] object[] objects,
            [ExcelArgument(Name = "Tag")] string tag = "",
            [ExcelArgument(Name = "Go")] bool go = false,
            [ExcelArgument(Name = "Return the pushed objects")] bool retObjs = false
            )
        {
            if (!go)
                return false;

            BHoMAdapter adapter = Project.ActiveProject.GetAdapter(adapterId);

            if (adapter == null)
                return "Failed to get adapter";

            List<IObject> iObjs = new List<IObject>();

            foreach (object obj in objects)
            {
                Guid guid;
                if (obj is string && Guid.TryParse(obj as string, out guid))
                {
                    IObject iOb = Project.ActiveProject.GetObject(guid);
                    if (iOb != null)
                        iObjs.Add(iOb);
                }
            }

            List<IObject> pushedObjects;

            try
            {
                pushedObjects = adapter.Push(iObjs, tag);
            }
            catch (Exception e)
            {
                return "Failed to push objects. Exception message: " + e.Message;
            }

            if (retObjs)
                return XlCall.Excel(XlCall.xlUDF, "Resize", pushedObjects.Select(x => x.ReturnTypeHelper()).ToArray());
            else
                return pushedObjects.Count == iObjs.Count;


        }

        /*****************************************************************/

        [ExcelFunction(Description = "Create an adapter", Category = "Dragon")]
        public static object Pull(
            [ExcelArgument(Name = "Adapter")] string adapterId,
            [ExcelArgument(Name = "Query")] string queryId,
            [ExcelArgument(Name = "Go")] bool go = false,
            [ExcelArgument(Name = "Return objects to list")] bool objsToList = false
            )
        {
            if (!go)
                return false;

            BHoMAdapter adapter = Project.ActiveProject.GetAdapter(adapterId);

            if (adapter == null)
                return "Failed to get adapter";

            IQuery query = Project.ActiveProject.GetQuery(queryId);

            if (adapter == null)
                return "Failed to get query";

            List<object> pulledObjs;

            try
            {
               pulledObjs = adapter.Pull(query).ToList();
            }
            catch (Exception e)
            {
                return "Failed to pull objects. Exception message: " + e.Message;
            }


            if (pulledObjs.Count < 1)
                return "No objects found";


            if (objsToList)
            {
                //TODO implement excel list....
                return Create.CreateList(pulledObjs.ToArray());
            }
            else
            {
                return XlCall.Excel(XlCall.xlUDF, "Resize", pulledObjs.Select(x => x.ReturnTypeHelper()).ToArray());
            }

        }

        /*****************************************************************/
    }
}
