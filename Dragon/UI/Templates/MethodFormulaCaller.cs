using BH.UI.Templates;
using ExcelDna.Integration;
using System;
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
                    params_ = "?by_"+ InputParams.Select(p => p.DataType.Name)
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

        public IEnumerable<ExcelArgumentAttribute> ExcelParams => Params.Select((p) => p.ArgumentAttribute);

        private void SetExcelMethod()
        {
            // Create an array of [n] parameters
            List<ParameterExpression> lambdaParams = new List<ParameterExpression>();
            foreach (var p in InputParams) {
                lambdaParams.Add(Expression.Parameter(typeof(object)));
            }
            Expression newArr = Expression.NewArrayInit(typeof(object), lambdaParams);
            MethodInfo storeMethod = DataAccessor.GetType().GetMethod("Store");
            MethodInfo returnMethod = DataAccessor.GetType().GetMethod("GetOutput");

            Expression accessorInstance = Expression.Convert(
                Expression.Constant(DataAccessor),
                typeof(FormulaDataAccessor)
            );
            Expression storeCall = Expression.Call(accessorInstance, storeMethod, newArr);
            Expression runCall = Expression.Call(Expression.Constant(this), this.GetType().GetMethod("Run", new Type[0]));
            Expression returnCall = Expression.Call(accessorInstance, returnMethod);
            Expression tree = Expression.Block(storeCall, runCall, returnCall);
            LambdaExpression lambda = Expression.Lambda(tree, lambdaParams);
            ExcelMethod = lambda.Compile();
        }
    }
}
