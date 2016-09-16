using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using BHG = BHoM.Geometry;
using BHB = BHoM.Base;

namespace Dragon
{
    public static class Functions
    {
        /*****************************************************************/
        /****  Geometry                                               ****/
        /*****************************************************************/

        [ExcelFunction(Description = "Create a BHoM point", Category = "Dragon")]
        public static object Point(double x, double y, double z)
        {
            return new BHG.Point(x, y, z).ToJSON();
        }

        /*****************************************************************/




    }
}
