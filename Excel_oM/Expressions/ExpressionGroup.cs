using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.oM.Excel.Expressions
{
    public class ExpressionGroup : IExpression
    {
        public IExpression Expression { get; set;  }
    }
}
