using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class PushFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter." + Caller.Name;

        public override Caller Caller { get; } = new PushCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public PushFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
