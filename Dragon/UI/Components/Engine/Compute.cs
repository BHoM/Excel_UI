using System;
using BH.oM.Base;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class ComputeFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ComputeCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ComputeFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
