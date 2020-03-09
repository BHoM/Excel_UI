using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.oM.Excel.Expressions
{
    public class BinaryExpression : IExpression
    {
        public string Operator { get; set; }
        public IExpression Left { get; set; }
        public IExpression Right { get; set; }
    }
}
