using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class FromJsonFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new FromJsonCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public FromJsonFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
