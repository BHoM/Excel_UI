using System;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;
using ExcelDna.Integration;
using System.Collections.Generic;

namespace BH.UI.Excel.Components
{
    public class PullFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter." + Caller.Name;

        public override Caller Caller { get; } = new PullCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public PullFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
