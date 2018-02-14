using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon
{
    public interface IExcelObject
    {
        Dictionary<string, object> PropertyDictionary();
        object InnerObject { get; }

        Guid BHoM_Guid { get; set; }
    }
}
