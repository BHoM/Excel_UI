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

        [ExcelFunction(Description = "Create a list of object", Category = "Dragon")]
        public static object CreateList(
                [ExcelArgument(Name = "Objects")] object[] objectsIds,
                [ExcelArgument(Name = "Type hint. Optional")] string typeHint = null
                )
        {

            if (objectsIds.Length < 1)
                return "No objects provided";
            List<object> objects = objectsIds.Select(x => x.CheckAndGetObjectOrGeometry()).ToList();


            bool sameType = true;
            Type t;

            //Check if a type hint has been provided
            if (string.IsNullOrWhiteSpace(typeHint))
            {
                //Check if all the objects are of the same type
                t = objects[0].GetType();

                for (int i = 1; i < objects.Count; i++)
                {
                    if (objects[i].GetType() != t)
                    {
                        sameType = false;
                        break;
                    }
                }
            }
            else
            {
                t = BH.Engine.Reflection.Create.Type(typeHint);
            }


            if (sameType)
            {
                var genericType = typeof(ExcelList<>);
                var specificType = genericType.MakeGenericType(t);
                var exlist = Activator.CreateInstance(specificType);

                PropertyInfo propData = specificType.GetProperty("Data");
                var data = propData.GetValue(exlist);

                MethodInfo add = data.GetType().GetMethod("Add");

                foreach (var item in objects)
                {
                    add.Invoke(data, new object[] { item });
                }

                Project.ActiveProject.AddBHoM(exlist as IObject);
                return (exlist as IObject).BHoM_Guid.ToString();
            }
            else
            {
                ExcelList<object> list = new ExcelList<object>();

                list.Data = objects;

                Project.ActiveProject.AddBHoM(list);
                return list.BHoM_Guid.ToString();
            }


        }

        /*****************************************************************/

        [ExcelFunction(Description = "Create a list of object", Category = "Dragon")]
        public static object CreateTuple(
        [ExcelArgument(Name = "Item 1")] object item1id,
        [ExcelArgument(Name = "Item 2")] object item2id,
        [ExcelArgument(Name = "Type hint1. Optional")] string typeHint1 = null,
        [ExcelArgument(Name = "Type hint2. Optional")] string typeHint2 = null
        )
        {
            object item1 = item1id.CheckAndGetObjectOrGeometry();
            object item2 = item2id.CheckAndGetObjectOrGeometry();

            Type t1, t2;

            t1 = string.IsNullOrWhiteSpace(typeHint1) ? item1.GetType() : BH.Engine.Reflection.Create.Type(typeHint1);
            t2 = string.IsNullOrWhiteSpace(typeHint2) ? item2.GetType() : BH.Engine.Reflection.Create.Type(typeHint2);

            var type = typeof(ExcelTuple<,>);
            var specificType = type.MakeGenericType(t1, t2);
            var exTuple = Activator.CreateInstance(specificType, item1, item2);

            Project.ActiveProject.AddBHoM(exTuple as IObject);
            return (exTuple as IObject).BHoM_Guid.ToString();
        }

        /*****************************************************************/

    }
}
