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

        public virtual string Name 
        {
            get
            {
                if (Caller is MethodCaller && Caller.SelectedItem != null)
                {
                    Type decltype = (Caller as MethodCaller).Method.DeclaringType;
                    return decltype.Name + "." + decltype.Namespace.Split('.').Last() + "." + Caller.Name;
                }
                return Caller.Category + "." + Caller.Name;
            }
        }
        public abstract Caller Caller { get; }
        public CallerFormula(FormulaDataAccessor accessor)
        {
            m_dataAccessor = accessor;
            Caller.SetDataAccessor(m_dataAccessor);
        }
    }
}
