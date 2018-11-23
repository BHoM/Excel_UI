using BH.Engine.Reflection;
using BH.Engine.Reflection.Convert;
using BH.UI.Dragon.Templates;
using BH.UI.Global;
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
        FormulaDataAccessor m_accessor;

        public FormulaSearchMenu(FormulaDataAccessor accessor)
        {
            m_accessor = accessor;
        }

        public override bool SetParent(object parent)
        {
            List<Delegate> delegates = new List<Delegate>();
            List<object> funcAttrs = new List<object>();
            List<List<object>> argAttrs = new List<List<object>>();
            foreach(var item in PossibleItems)
            {
                try
                {
                    var proxy = CreateDelegate(item.Key, item.Value);
                    delegates.Add(proxy.Item1);
                    funcAttrs.Add(proxy.Item2);
                    argAttrs.Add(proxy.Item3);
                } catch (Exception e) {
                    Console.WriteLine(e.Message);
                }
            }
            try
            {
                ExcelIntegration.RegisterDelegates(delegates, funcAttrs, argAttrs);
            } catch
            {
                return false;
            }
            return true;
        }

        private Tuple<Delegate, ExcelFunctionAttribute, List<object>> CreateDelegate(string itemStr, MethodInfo item)
        {
            return m_accessor.Wrap(item, () => NotifySelection(itemStr));
        }

        private ParameterExpression[] GetParams(IEnumerable<ParameterInfo> input)
        {
            return input.Select(p => Expression.Parameter(typeof(object))).ToArray();
        }
    }
}
