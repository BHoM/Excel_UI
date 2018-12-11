using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class GetPropertyFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new GetPropertyCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public GetPropertyFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
