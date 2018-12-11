using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class QueryFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new QueryCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public QueryFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
