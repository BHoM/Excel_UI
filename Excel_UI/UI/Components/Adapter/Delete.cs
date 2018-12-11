using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class DeleteFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter." + Caller.Name;

        public override Caller Caller { get; } = new DeleteCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public DeleteFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
