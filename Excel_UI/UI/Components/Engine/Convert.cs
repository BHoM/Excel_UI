using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class ConvertFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ConvertCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ConvertFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
