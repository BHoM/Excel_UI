using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.oM.Excel.Expressions
{
    public class UnaryExpression : IExpression
    {
        public string Operator { get; set; }
        public IExpression Expression { get; set; }
    }
}
