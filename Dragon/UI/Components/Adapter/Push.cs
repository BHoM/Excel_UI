using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class PushFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter." + Caller.Name;

        public override Caller Caller { get; } = new PushCaller();

        public PushFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
