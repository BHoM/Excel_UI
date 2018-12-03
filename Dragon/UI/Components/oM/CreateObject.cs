using BH.UI.Components;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.Components
{
    public class CreateObjectFormula : CallerFormula
    {

        public override Caller Caller { get; } = new CreateObjectCaller();

        public CreateObjectFormula(FormulaDataAccessor accessor) : base(accessor) { }
    }
}
