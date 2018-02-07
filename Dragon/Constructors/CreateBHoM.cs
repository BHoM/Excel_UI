using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using BH.oM.Base;
using BH.Engine.Reflection;

namespace BH.UI.Dragon
{
    public static partial class Create
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM object", Category = "Dragon")]
        public static object CreateObject(
            [ExcelArgument(Name = "object type")] string typeString,
            [ExcelArgument(Name = "property names (optional)")] object[] propNames,
            [ExcelArgument(Name = "property values (optional)")] object[] propValues)
        {

            if (propNames.Length != propValues.Length)
                return "Need to provide the same number of property names as property values";

            Type type = BH.Engine.Reflection.Create.Type(typeString);
            BHoMObject obj = type.GetConstructor(Type.EmptyTypes).Invoke(new object[] { }) as BHoMObject;

            string message;
            if (!InOutHelp.SetPropertyHelper(obj, propNames, propValues, out message))
                return message;

            Project.ActiveProject.AddObject(obj);
            return obj.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Create a CustomObject", Category = "Dragon")]
        public static object CreateCustomObject(
            [ExcelArgument(Name = "property names (optional)")] object[] propNames,
            [ExcelArgument(Name = "property values (optional)")] object[] propValues)
        {

            if (propNames.Length != propValues.Length)
                return "Need to provide the same number of property names as property values";

            CustomObject obj = new CustomObject();

            string message;
            if (!InOutHelp.SetPropertyHelper(obj, propNames, propValues, out message))
                return message;

            Project.ActiveProject.AddObject(obj);
            return obj.BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Call the ToString() method from an object", Category = "Dragon")]
        public static object ToString(
            [ExcelArgument(Name = "object id")] string objectId)
        {
            IObject obj = Project.ActiveProject.GetObject(objectId);

            return obj.ToString();
        }

        /*****************************************************************/
        [ExcelFunction(Description = "Get the names of all the types of BHoMObjects", Category = "Dragon")]
        public static object GetAllObjectTypes(
            [ExcelArgument(Name = "Vertical expansion of names")] bool vertical = true)
        {
            List<string> objectNames = new List<string>();

            foreach (KeyValuePair<string, List<Type>> kvp in Query.BHoMTypeDictionary())
            {
                objectNames.Add(kvp.Key);
            }


            return XlCall.Excel(XlCall.xlUDF, "Resize", objectNames.ToArray());
        }

        /*****************************************************************/

        //[ExcelFunction(Description = "Get a definition of all the BhoM objects", Category = "Dragon")]
        //public static object GetAllObjectModels()
        //{

        //    List<string[]> data = new List<string[]>();
        //    foreach (KeyValuePair<string, Type> kvp in  Query.GetBHoMTypeDictionary())
        //    {
        //        if (!kvp.Key.Contains('.')) continue; // Need a better way to access each type only once

        //        string[] trow = new string[3];
        //        trow[0] = kvp.Key;
        //        trow[1] = "";
        //        trow[2] = "";
        //        data.Add(trow);

        //        foreach (PropertyInfo prop in kvp.Value.GetProperties())
        //        {
        //            if (prop.CanRead && prop.CanWrite)
        //            {
        //                string[] row = new string[3];
        //                row[0] = "";
        //                row[1] = prop.Name;
        //                row[2] = prop.PropertyType.ToString();
        //                data.Add(row);
        //            }
        //        }
        //    }

        //    int nb = data.Count;
        //    object[,] array = new object[nb, 3];
        //    for (int i = 0; i < nb; i++)
        //    {
        //        for (int j = 0; j < 3; j++)
        //            array[i, j] = data[i][j];
        //    }

        //    return XlCall.Excel(XlCall.xlUDF, "Resize", array);
        //}


    }
}
