using System;

using BH.oM.Base;

using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class MoveFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter." + Caller.Name;

        public override Caller Caller { get; } = new MoveCaller();

        public MoveFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
