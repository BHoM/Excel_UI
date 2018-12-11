using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class ToJsonFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ToJsonCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ToJsonFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
