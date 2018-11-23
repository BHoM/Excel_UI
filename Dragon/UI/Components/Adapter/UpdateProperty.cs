using System;
using BH.oM.Base;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class UpdatePropertyFormula : SingleMethodCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new UpdatePropertyCaller();

        public override string Name => "Adapter.UpdateProperty";

        public override string Category => "Dragon.Adapter";

        public UpdatePropertyFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
