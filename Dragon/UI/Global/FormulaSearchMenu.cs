using BH.Engine.Reflection;
using BH.Engine.Reflection.Convert;
using BH.oM.UI;
using BH.UI.Dragon.Templates;
using BH.UI.Global;
using BH.UI.Templates;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.Global
{
    public class FormulaSearchMenu : SearchMenu
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public FormulaSearchMenu(FormulaDataAccessor accessor, Dictionary<string, CallerFormula> callers) : base()
        {
            m_accessor = accessor;
            m_callers = callers;
        }

        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        public override bool SetParent(object parent)
        {
            List<Delegate> delegates = new List<Delegate>();
            List<ExcelFunctionAttribute> funcAttrs = new List<ExcelFunctionAttribute>();
            List<List<object>> argAttrs = new List<List<object>>();
            Dictionary<string, int> dups = new Dictionary<string, int>();
            foreach(var item in PossibleItems)
            {
                try
                {
                    var proxy = CreateDelegate(item);
                    if (proxy == null) continue;
                    var name = proxy.Item2.Name;
                    if (!dups.ContainsKey(name))
                    {
                        dups.Add(name, 1);
                        delegates.Add(proxy.Item1);
                        funcAttrs.Add(proxy.Item2);
                        argAttrs.Add(proxy.Item3);
                    }
                } catch (Exception e) {
                    Console.WriteLine(e.Message);
                }
            }
            try
            {
                ExcelIntegration.RegisterDelegates(delegates, funcAttrs.Cast<object>().ToList(), argAttrs);
            } catch
            {
                return false;
            }
            return true;
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private Tuple<Delegate, ExcelFunctionAttribute, List<object>> CreateDelegate(SearchItem item)
        {
            if (m_callers.ContainsKey(item.CallerType.Name))
            {
                CallerFormula caller = m_callers[item.CallerType.Name];
                caller.Caller.SetItem(item.Item);
                return m_accessor.Wrap(caller, () => NotifySelection(item));
            }
            return null;
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private FormulaDataAccessor m_accessor;
        private Dictionary<string, CallerFormula> m_callers;
    }
}
