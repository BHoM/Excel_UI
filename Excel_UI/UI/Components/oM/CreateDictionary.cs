using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Templates;
using BH.UI.Components;

namespace BH.UI.Excel.Components
{
    public class CreateDictionaryFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new CreateDictionaryCaller();

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/


        public CreateDictionaryFormula(FormulaDataAccessor accessor) : base(accessor) { }
        /*******************************************/
    }
}
