using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Structural.Loads;
using BH.oM.Structural.Elements;
using BH.oM.Geometry;
using BH.oM.Structural.Properties;

namespace BH.UI.Dragon
{
    public static class Tests
    {

        public static string TestEnumDefault(LoadNature nature = LoadNature.Dead)
        {
            return nature.ToString();
        }

        public static Node TestNode(Point point = null)
        {
            return point != null ? new Node() { Position = point } : new Node() { Position = new Point() { X = 1, Y = 2, Z = 3 } };
        }
    }
}
