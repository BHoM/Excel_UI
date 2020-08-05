using BH.Engine.Excel;
using BH.Engine.Base;
using BH.oM.Excel.Expressions;
using BH.oM.UI;
using BH.UI.Excel.Templates;
using BH.UI.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace BH.UI.Excel.UI.Global
{
    class ComponentUpgrader
    {
        /*************************************/
        /**** Constructors                ****/
        /*************************************/

        public ComponentUpgrader(string oldFormula, CallerFormula caller)
        {
            m_NewName = caller.Function;
            m_NewParams = caller.Caller.InputParams.ToList();
            m_OldName = oldFormula;
            m_Upgraded = caller.Caller.WasUpgraded;
            Register();
        }

        /*************************************/
        /**** Private Methods             ****/
        /*************************************/

        private void Register()
        {
            lock (m_Mutex)
            {
                if (m_Registered.Contains(m_OldName))
                    return;
                ExcelIntegration.RegisterDelegates(
                    new List<Delegate> { new Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object, object>((a, b, c, d, e, f, g, h, i, j, k, l, m, n) => Upgrade()) },
                    new List<object> { new ExcelFunctionAttribute { Name = m_OldName } },
                    new List<List<object>> { new List<object> { } }
                );
            }
        }

        /*************************************/

        private object Upgrade()
        {
            ExcelReference xlref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                string formula = XlCall.Excel(XlCall.xlfGetFormula, xlref).ToString();
                var expr = formula.ToExpression();
                var newExpr = IRemap(expr);
                xlref.ToReference().Contents("=" + newExpr.IToFormula());
            });
            return ExcelError.ExcelErrorName;
        }

        /*************************************/

        private IExpression IRemap(IExpression expression)
        {
            return Remap(expression as dynamic);
        }

        /*************************************/

        private IExpression Remap(FunctionExpression expression)
        {
            if (expression.Name == m_OldName)
            {
                var newExpr = new FunctionExpression { Name = m_NewName };
                if (!m_Upgraded) // If just renamed
                {
                    newExpr.Arguments = expression.Arguments.Select(IRemap).ToList();
                }
                else
                {
                    foreach (var param in m_NewParams)
                    {
                        IExpression newParam = new EmptyExpression();
                        var oldIndexFgmt = param.FindFragment<ParamOldIndexFragment>();
                        int oldIndex = -1;
                        if (oldIndexFgmt != null)
                        {
                            oldIndex = oldIndexFgmt.OldIndex;
                        }
                        if (oldIndex != -1 && oldIndex < expression.Arguments.Count)
                        {
                            newParam = IRemap(expression.Arguments[oldIndex]);
                        }
                        newExpr.Arguments.Add(newParam);
                    }
                }

                // Remove trailing empty expressions
                int lastNonEmpty = newExpr.Arguments.FindLastIndex(e => !(e is EmptyExpression)) + 1;
                if (lastNonEmpty < newExpr.Arguments.Count)
                    newExpr.Arguments.RemoveRange(lastNonEmpty, newExpr.Arguments.Count - lastNonEmpty);
                return newExpr;
            }
            return new FunctionExpression { Name = expression.Name, Arguments = expression.Arguments.Select(IRemap).ToList() };
        }

        /*************************************/

        private IExpression Remap(ArrayExpression expression)
        {
            return new ArrayExpression { Expressions = expression.Expressions.Select(IRemap).ToList() };
        }

        /*************************************/

        private IExpression Remap(ExpressionGroup expression)
        {
            return new ExpressionGroup { Expression = IRemap(expression.Expression) };
        }

        /*************************************/

        private IExpression Remap(BinaryExpression expression)
        {
            return new BinaryExpression { Operator = expression.Operator, Left = IRemap(expression.Left), Right = IRemap(expression.Right) };
        }

        /*************************************/

        private IExpression Remap(UnaryExpression expression)
        {
            return new UnaryExpression { Operator = expression.Operator, Expression = IRemap(expression.Expression) };
        }

        /*************************************/

        private IExpression Remap(IExpression expression)
        {
            return expression;
        }

        /*************************************/
        /**** Private Fields              ****/
        /*************************************/

        private string m_NewName;
        private List<ParamInfo> m_NewParams;
        private string m_OldName;
        private bool m_Upgraded;
        private static HashSet<string> m_Registered = new HashSet<string>();
        private static object m_Mutex = new object();

        /*************************************/
    }
}
