using System;
using BH.oM.Base;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class DeleteFormula : SingleMethodCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new DeleteCaller();

        public override string Category => "Dragon.Adapter";

        public override string Name => "Adapter.Delete";

        public DeleteFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
