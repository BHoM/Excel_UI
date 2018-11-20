using BH.UI.Templates;
using ExcelDna.Integration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.UI.Templates
{
    public class MethodFormulaCaller : MethodCaller, IFormulaCaller
    {
        public MethodFormulaCaller(MethodBase method) : base(method)
        {
            Params = InputParams.Select(p => new FormulaParameter(p));
            SetDataAccessor(new FormulaDataAccessor(InputParams));
            SetExcelMethod();
        }

        public Delegate ExcelMethod { get; protected set; }

        public ExcelFunctionAttribute FunctionAttribute
        {
            get
            {
                bool hasParams = InputParams.Count() > 0;
                string params_ = "";
                if (hasParams) {
                    params_ = "?by_" + InputParams
                        .Select(p => p.Name )
                        .Aggregate((a, b) => $"{a}_{b}");
                }
                return new ExcelFunctionAttribute()
                {
                
                    Name = Method.DeclaringType.Name+"."
                        +Method.DeclaringType.Namespace.Split('.').Last()
                        +"."+Name + params_,
                    Description = Description,
                    Category = "Dragon."+Category
                };
            }
        }

        public IEnumerable<IFormulaParameter> Params { get; protected set; }

        public IEnumerable<ExcelArgumentAttribute> ExcelParams =>
            Params.Select((p) => p.ArgumentAttribute);

        private void SetExcelMethod()
        {
            // Create a Delegate that looks like:
            //
            // (a, b, c, ...) => {
            //     DataAccessor.Store( new [] {a, b, c, ...} );
            //     Run();
            //     return DataAccessor.GetOutput();
            // }

            // Create an array of [n] parameters
            List<ParameterExpression> lambdaParams =
                new List<ParameterExpression>();
            foreach (var p in InputParams) {
                lambdaParams.Add(Expression.Parameter(typeof(object)));
            }
            Expression newArr = Expression.NewArrayInit(
                typeof(object), // new []
                lambdaParams    // { ... }
            );

            Type accessorType = typeof(FormulaDataAccessor);

            // Convert DataAccessor to FormulaDataAccessor
            Expression accessorInstance = Expression.Convert(
                Expression.Constant(DataAccessor),
                accessorType
            );

            // Call FormulaDataAccessor.Store with array
            MethodInfo storeMethod = accessorType.GetMethod("Store");
            Expression storeCall = Expression.Call(
                accessorInstance, // (FormulaDataAccessor)DataAccessor
                storeMethod,      // void Store(object[])
                newArr            // new [] { ... }
            );

            // Call this.Run()
            MethodInfo runMethod = GetType().GetMethod("Run", new Type[0]);
            Expression runCall = Expression.Call(
                Expression.Constant(this),
                runMethod
            );

            // Return call FormulaDataAccessor.GetOutput()
            MethodInfo returnMethod = accessorType.GetMethod("GetOutput");
            Expression returnCall = Expression.Call(
                accessorInstance, // (FormulaDataAccessor)DataAccessor
                returnMethod      // object GetOutput()
            );

            // Chain them together
            Expression tree = Expression.Block(storeCall, runCall, returnCall);
            LambdaExpression lambda = Expression.Lambda(tree, lambdaParams);

            // Compile
            ExcelMethod = lambda.Compile();
        }
    }
}
