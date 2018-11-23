using BH.UI.Templates;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.Templates
{
    public abstract class CallerFormula
    {
        private FormulaDataAccessor m_dataAccessor;

        public abstract Caller Caller { get; }
        public CallerFormula(FormulaDataAccessor accessor)
        {
            m_dataAccessor = accessor;
            Caller.SetDataAccessor(m_dataAccessor);
            Caller.ItemSelected += OnItemSelected;
        }

        private void OnItemSelected(object sender, object e)
        {
        }
    }
}
