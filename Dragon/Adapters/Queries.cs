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


namespace BH.UI.Dragon.Adapters
{
    public static class Queries
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/

        [ExcelFunction(Description = "Get available adapters", Category = "Dragon")]
        public static object CreateFilterQuery(
            [ExcelArgument(Name = "Type")] string typeString,
            [ExcelArgument(Name = "Tag")] string tag = "",
            [ExcelArgument(Name = "Equalities Names. Optional")] object[] eqName = null,
            [ExcelArgument(Name = "Equalities Values. Optional")] object[] eqVal = null)
        {
            Type type;

            List<Type> types;

            if (Query.BHoMTypeDictionary().TryGetValue(typeString, out types))
            {
                if (types.Count > 1)
                    return "Mutliple types found with the given name. Please be more specific using leading namespaces. Example: BH.oM.Structural.Elements.Bar instead of Bar";
                else
                    type = types[0];
            }
            else
                return "No type found mathing the input. Please check spelling";

            Dictionary<string, object> equalities = new Dictionary<string, object>();

            if (InOutHelp.IsValidArray(eqName) && InOutHelp.IsValidArray(eqVal))
            {
                if (eqName.Length != eqVal.Length)
                    return "Need same number of Equalities names as equalities values. Currently provided " + eqName.Length + " names and " + eqVal.Length + " values.";

                for (int i = 0; i < eqName.Length; i++)
                {
                    equalities[eqName[i] as string] = eqVal[i].CheckAndGetObjectOrGeometry();
                }
            }

            FilterQuery query = new FilterQuery() { Tag = tag, Type = type, Equalities = equalities };

            return Project.ActiveProject.AddQuery(query).ToString();
        }

        /*****************************************************************/

    }
}
