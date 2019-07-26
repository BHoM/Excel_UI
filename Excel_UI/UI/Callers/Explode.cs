using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Callers
{
    class ExplodeCaller : UI.Templates.MethodCaller
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ExplodeCaller() : base(typeof(Properties).GetMethod("Explode"))
        {
        }
    }
}
