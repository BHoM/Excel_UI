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
            List<object> objects = objectsIds.Select(x => x.CheckAndGetStoredObject()).ToList();


            bool sameType;
            Type t = GetTypeFromObjectsAndHint(objects, typeHint, out sameType);


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

                return Project.ActiveProject.Add(exlist as IExcelObject).ToString();
            }
            else
            {
                ExcelList<object> list = new ExcelList<object>();

                list.Data = objects;

                Project.ActiveProject.Add(list);
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
            object item1 = item1id.CheckAndGetStoredObject();
            object item2 = item2id.CheckAndGetStoredObject();

            Type t1, t2;

            t1 = string.IsNullOrWhiteSpace(typeHint1) ? item1.GetType() : BH.Engine.Reflection.Create.Type(typeHint1);
            t2 = string.IsNullOrWhiteSpace(typeHint2) ? item2.GetType() : BH.Engine.Reflection.Create.Type(typeHint2);

            var type = typeof(ExcelTuple<,>);
            var specificType = type.MakeGenericType(t1, t2);
            var exTuple = Activator.CreateInstance(specificType, item1, item2);

            Project.ActiveProject.Add(exTuple as IBHoMObject);
            return (exTuple as IBHoMObject).BHoM_Guid.ToString();
        }

        /*****************************************************************/

        [ExcelFunction(Description = "Create a Dictionary", Category = "Dragon")]
        public static object CreateDictionary(
                [ExcelArgument(Name = "Keys")] object[] keysIds,
                [ExcelArgument(Name = "Values")] object[] valuesIds,
                [ExcelArgument(Name = "Type hint keys. Optional")] string typeHintKey = null,
                [ExcelArgument(Name = "Type hint values. Optional")] string typeHintValue = null
                )
        {

            if (keysIds.Length != valuesIds.Length)
                return "Need to provide the same number of keys as values";

            if (keysIds.Length < 1)
                return "No objects provided";

            List<object> keys = keysIds.Select(x => x.CheckAndGetStoredObject()).ToList();
            List<object> values = valuesIds.Select(x => x.CheckAndGetStoredObject()).ToList();


            bool sameTypeKey, sameTypeValue;

            Type keyType = GetTypeFromObjectsAndHint(keys, typeHintKey, out sameTypeKey);
            Type valueType = GetTypeFromObjectsAndHint(values, typeHintValue, out sameTypeValue);



            Type dictionaryType = typeof(ExcelDictionary<,>).MakeGenericType(new Type[] { keyType, valueType });
            var dictionary = Activator.CreateInstance(dictionaryType);

            PropertyInfo propData = dictionaryType.GetProperty("Data");
            var data = propData.GetValue(dictionary);
            MethodInfo add = data.GetType().GetMethod("Add");

            for (int i = 0; i < keys.Count; i++)
            {
                add.Invoke(data, new object[] { keys[i], values[i] });
            }

            return Project.ActiveProject.Add(dictionary as IExcelObject).ToString();


        }

        /*****************************************************************/
        /******* Private methods                            **************/
        /*****************************************************************/

        private static Type GetTypeFromObjectsAndHint(List<object> objects, string typeHint, out bool sameType)
        {
            sameType = true;
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

                if (!sameType)
                    t = typeof(object);
            }
            else
            {
                t = BH.Engine.Reflection.Create.Type(typeHint);
            }

            return t;
        }

        /*****************************************************************/

    }
}
