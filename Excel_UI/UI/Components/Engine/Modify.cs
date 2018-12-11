using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class ModifyFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ModifyCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ModifyFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
