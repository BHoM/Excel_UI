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

        public ExplodeCaller() : base(typeof(Methods.Properties).GetMethod("Explode"))
        {
        }

        public override System.Drawing.Bitmap Icon_24x24 => m_native.Icon_24x24;

        public override string Name => m_native.Name;

        public override string Category => m_native.Category;

        public override string Description => m_native.Description;

        public override int GroupIndex => m_native.GroupIndex;

        private UI.Components.ExplodeCaller m_native = new UI.Components.ExplodeCaller();
    }
}
