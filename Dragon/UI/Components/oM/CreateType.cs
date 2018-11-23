using System;
using BH.oM.Base;
using BH.UI.Dragon.Templates;
using BH.UI.Templates;
using BH.UI.Components;
using BH.Engine.Reflection.Convert;

namespace BH.UI.Dragon.Components
{
    public class CreateTypeFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new CreateTypeCaller();


        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CreateTypeFormula(FormulaDataAccessor accessor) : base(accessor)
        {
            Caller.ItemSelected += DynamicCaller_ItemSelected;
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private void DynamicCaller_ItemSelected(object sender, object e)
        {
            Type type = e as Type;

            //if (type != null)
                //Message = type.ToText();
        }


        /*******************************************/
    }
}
