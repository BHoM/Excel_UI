using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.oM.Excel.Expressions
{
    public class FunctionExpression : IExpression
    {
        public string Name { get; set; }
        public List<IExpression> Arguments { get; set; } = new List<IExpression>();
    }
}
