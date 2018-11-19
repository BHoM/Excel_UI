using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.UI.Templates
{
    public interface IFormulaParameter
    {
        ExcelArgumentAttribute ArgumentAttribute { get; }
    }
}
