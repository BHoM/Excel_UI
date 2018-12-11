using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class ExplodeFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ExplodeCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ExplodeFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
