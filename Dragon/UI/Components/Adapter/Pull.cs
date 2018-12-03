using System;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;
using ExcelDna.Integration;
using System.Collections.Generic;

namespace BH.UI.Dragon.Components
{
    public class PullFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter." + Caller.Name;

        /*******************************************/

        public override Caller Caller { get; } = new PullCaller();

        /*******************************************/

        public PullFormula(FormulaDataAccessor accessor) : base(accessor)
        {
        }

        /*******************************************/
    }
}
