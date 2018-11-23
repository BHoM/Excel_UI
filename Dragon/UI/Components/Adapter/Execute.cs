using System;

using BH.oM.Base;

using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class ExecuteFormula : SingleMethodCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new ExecuteCaller();

        public override string Category => "Dragon.Adapter";

        public override string Name => "Adapter.Execute";

        public ExecuteFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
