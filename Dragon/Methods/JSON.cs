using ExcelDna.Integration;
using BH.Engine.Serialiser;

namespace BH.UI.Dragon
{
    public static class JSON
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get the Json string of an object", Category = "Dragon")]
        public static object ToJSON(
            [ExcelArgument(Name = "object id")] string objectId)
        {

            //Get out the object
            object obj = Project.ActiveProject.GetAny(objectId);

            if (obj == null)
                return "Object not found";

            return obj.ToJson();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Get the Json string of an object", Category = "Dragon")]
        public static object FromJSON(
                [ExcelArgument(Name = "object id")] string json)
        {

            //Get out the object
            object obj = Convert.FromJson(json);

            if (obj == null)
                return "Object failed to serialize";

            return obj.ReturnTypeHelper();
        }

        /*****************************************************************/
    }
}
