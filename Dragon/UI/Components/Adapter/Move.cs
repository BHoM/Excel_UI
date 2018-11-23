using System;

using BH.oM.Base;

using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class MoveFormula : SingleMethodCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new MoveCaller();

        public override string Category => "Dragon.Adapter";

        public override string Name => "Adapter.Move";

        public MoveFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
