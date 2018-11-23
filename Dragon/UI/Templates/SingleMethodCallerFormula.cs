using BH.UI.Templates;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.Templates
{
    public abstract class SingleMethodCallerFormula : CallerFormula
    {
        public abstract string Name { get; }
        public abstract string Category { get; }

        public SingleMethodCallerFormula(FormulaDataAccessor accessor): base(accessor)
        {
            var methodCaller = Caller as MethodCaller;
            if (methodCaller == null) return;
            var proxy = accessor.Wrap(methodCaller.Method, () => Caller.Run());
            ExcelIntegration.RegisterDelegates(
                new List<Delegate>() { proxy.Item1 },
                new List<object>()
                {
                    new ExcelFunctionAttribute()
                    {
                        Name = Name,
                        Description = methodCaller.Description,
                        Category = Category
                    }
                },
                new List<List<object>>()
                {
                    proxy.Item3
                }
            );
        }
    }
}
