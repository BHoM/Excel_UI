using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Dragon.Components
{
    public class PushFormula : SingleMethodCallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new PushCaller();

        public override string Name => "Adapter.Push";

        public override string Category => "Dragon.Adapter";

        public PushFormula(FormulaDataAccessor accessor) : base(accessor) { }

        /*******************************************/
    }
}
