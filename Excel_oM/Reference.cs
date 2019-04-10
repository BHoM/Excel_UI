using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.oM.Excel
{
    public class Reference : BHoMObject
    {
        public List<Rectangle> Rectangles { get; set; }
        public IntPtr Sheet { get; set; }
    }
}
