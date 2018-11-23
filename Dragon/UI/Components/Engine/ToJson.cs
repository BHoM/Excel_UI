using System;
using BH.oM.Base;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class ToJsonFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ToJsonCaller();

        public ToJsonFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
