using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class CreateQueryFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter.Query." + Caller.Name;

        public override Caller Caller { get; } = new CreateQueryCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CreateQueryFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
