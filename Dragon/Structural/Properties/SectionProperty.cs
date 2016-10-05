using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BHP = BHoM.Structural.Properties;
using ExcelDna.Integration;
using BHB = BHoM.Base;
using BHG = BHoM.Global;
using System.Reflection;

namespace Dragon.Structural.Properties
{
    public static class SectionProperty
    {
        [ExcelFunction(Description = "Load a steel section from DB", Category = "Dragon.Structural")]
        public static object SteelSectionFromDB(
            [ExcelArgument(Name = "Section name")] string name)
        {
            BHP.SectionProperty prop = BHP.SectionProperty.LoadFromSteelSectionDB(name);

            BHG.Project.ActiveProject.AddObject(prop);
            return prop.BHoM_Guid.ToString();
        }

        [ExcelFunction(Description = "Load a cable section from DB based on diameter", Category = "Dragon.Structural")]
        public static object CableSectionFromDBDiameter(
        [ExcelArgument(Name = "Section diameter in [mm]")] double diameter,
        [ExcelArgument(Name = "Number of cables in the section")] int nb = 1)
        {
            diameter = diameter / 1000;
            BHP.SectionProperty prop = BHP.SectionProperty.LoadFromCableSectionDBDiameter(diameter, nb);

            BHG.Project.ActiveProject.AddObject(prop);
            return prop.BHoM_Guid.ToString();
        }

        [ExcelFunction(Description = "Load a cable section from DB based on name", Category = "Dragon.Structural")]
        public static object CableSectionFromDBName(
        [ExcelArgument(Name = "Section name")] string name,
        [ExcelArgument(Name = "Number of cables in the section")] int nb = 1)
        {
            BHP.SectionProperty prop = BHP.SectionProperty.LoadFromCableSectionDBName(name, nb);

            BHG.Project.ActiveProject.AddObject(prop);
            return prop.BHoM_Guid.ToString();
        }

        [ExcelFunction(Description = "Load a steel section from DB", Category = "Dragon.Structural")]
        public static object CreateSteelSectionFromString(
            [ExcelArgument(Name = "Section properties")] string name)
        {
            BHP.SectionProperty prop = BHP.SectionProperty.LoadFromSteelSectionDB(name);

            if(prop == null)
                prop = BHP.SectionProperty.CreateSectionPropertyFromString(name);

            if (prop != null)
            {
                BHG.Project.ActiveProject.AddObject(prop);
                return prop.BHoM_Guid.ToString();
            }
            return "Creation Failed";
        }
    }
}
