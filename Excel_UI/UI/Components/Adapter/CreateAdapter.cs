using BH.UI.Components;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Components
{
    public class CreateAdapterFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override string Name => "Adapter.Create." + Caller.Name;

        public override Caller Caller { get; } = new CreateAdapterCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CreateAdapterFormula(FormulaDataAccessor accessor) : base(accessor) {
        }
    }
}
