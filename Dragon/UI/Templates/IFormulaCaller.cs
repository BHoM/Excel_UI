using BH.UI.Templates;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.UI.Templates
{
    public interface IFormulaCaller 
    {
        Delegate ExcelMethod { get; }
        ExcelFunctionAttribute FunctionAttribute { get; }
        IEnumerable<IFormulaParameter> Params { get; }
        IEnumerable<ExcelArgumentAttribute> ExcelParams { get; }
    }
}
