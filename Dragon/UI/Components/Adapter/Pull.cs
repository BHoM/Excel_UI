using System;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;
using ExcelDna.Integration;
using System.Collections.Generic;

namespace BH.UI.Dragon.Components
{
    public class PullFormula : SingleMethodCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new PullCaller();

        public override string Name => "Adapter.Pull";

        public override string Category => "Dragon.Adapter";

        public PullFormula(FormulaDataAccessor accessor) : base(accessor)
        {
        }

        /*******************************************/
    }
}
