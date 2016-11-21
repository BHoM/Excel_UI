using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BHL = BHoM.Structural.Loads;
using ExcelDna.Integration;
using BHB = BHoM.Base;
using BHG = BHoM.Global;
using BGE = BHoM.Geometry;
using BHE = BHoM.Structural.Elements;
using System.Reflection;


namespace Dragon.Structural.Loads
{
    public static class Loads
    {
        [ExcelFunction(Description = "Create a load case", Category = "Dragon.Structural")]
        public static object CreateLoadCase(
        [ExcelArgument(Name = "Name of the load case")] object name,
        [ExcelArgument(Name = "Nature of the case")]  object nature,
        [ExcelArgument(Name = "Load case number")]  int caseNo)
        {
            BHL.LoadNature nat;

            if (!Enum.TryParse(nature.ToString(), out nat))
                return "Operation failed. Try changing load nature name";
            BHL.Loadcase loadCase = new BHL.Loadcase(name.ToString(), nat);
            loadCase.Number = caseNo;

            BHG.Project.ActiveProject.AddObject(loadCase);
            return loadCase.BHoM_Guid.ToString();
        }

        [ExcelFunction(Description = "Create a Load combination", Category = "Dragon.Structural")]
        public static object CreateLoadCombination(
            [ExcelArgument(Name = "Name of the load case")] string name,
            [ExcelArgument(Name = "Load Cases")]  object[] caseIds,
            [ExcelArgument(Name = "Load Factors")]  object[] loadFactors)
        {

            List<BHL.ICase> cases = new List<BHL.ICase>();
            List<double> factors = new List<double>();

            if (caseIds.Length != loadFactors.Length)
                return "Need to provide the same number of load factors as cases";

            for (int i = 0; i < caseIds.Length; i++)
            {
                double d;
                if (!double.TryParse(loadFactors[i].ToString(), out d))
                    continue;

                if (d <= 0)
                    continue;

                object caseObj = BHG.Project.ActiveProject.GetObject(caseIds[i].ToString());

                BHL.ICase loadCase = (BHL.ICase)caseObj;

                if (loadCase == null)
                    continue;

                cases.Add(loadCase);
                factors.Add(d);
            }

            if (cases.Count < 1)
                return "";

            BHL.LoadCombination comb = new BHL.LoadCombination(name, cases, factors);

            BHG.Project.ActiveProject.AddObject(comb);
            return comb.BHoM_Guid.ToString();
        }

        [ExcelFunction(Description = "Create a Load combination. Assumed units are [kN]/[kNm]/[m]/[°C]", Category = "Dragon.Structural")]
        public static object CreateLoad(
        [ExcelArgument(Name = "Load Case")]  object caseId,
        [ExcelArgument(Name = "Load Type")]  string loadType,
        [ExcelArgument(Name = "Magnitude x-y-z-mx-my-mz")]  double[] magnitude,
        [ExcelArgument(Name = "Group name")]  string groupName,
        [ExcelArgument(Name = "Global or local axis. N=Global, Y = Local")]  string axis = null,
        [ExcelArgument(Name = "Global or local axis. Y = yes, N = No")]  string projected = null)
        {

            double sFac = 1000;

            BHL.Loadcase loadCase = (BHL.Loadcase)BHG.Project.ActiveProject.GetObject(caseId.ToString());
            BHB.IGroup group;

            BGE.Vector force = null;
            BGE.Vector moment = null;
            double mag;

            if (magnitude.Length < 1)
                return "Load needs magnitude";

            mag = magnitude[0];

            if (magnitude.Length > 2)
                force = new BGE.Vector(magnitude[0], magnitude[1], magnitude[2]);

            if (magnitude.Length > 5)
                moment = new BGE.Vector(magnitude[3], magnitude[4], magnitude[5]);

            BHL.ILoad load;

            switch (loadType)
            {

                case "BarUDL":
                    if (force == null)
                        return "BarUDL load needs load vector";

                    BHL.BarUniformlyDistributedLoad udlLoad = new BHL.BarUniformlyDistributedLoad();
                    group = new BHB.Group<BHE.Bar>();
                    group.Name = groupName;
                    udlLoad.Objects = (BHB.Group<BHE.Bar>)group;
                    udlLoad.ForceVector = force* sFac;

                    if (moment != null)
                        udlLoad.MomentVector = moment* sFac;
                    load = udlLoad;
                    break;
                case "NodeDisplacement":
                    if (force == null)
                        return "Node force needs force vector";
                    BHL.PointDisplacement ptDisp = new BHL.PointDisplacement();
                    group = new BHB.Group<BHE.Node>();
                    group.Name = groupName;
                    ptDisp.Objects = (BHB.Group<BHE.Node>)group;
                    ptDisp.Translation = force;

                    if (moment != null)
                        ptDisp.Rotation = moment;
                    load = ptDisp;
                    break;
                case "NodeForce":
                    if (force == null)
                        return "Node force needs force vector";
                    BHL.PointForce ptForce = new BHoM.Structural.Loads.PointForce();
                    group = new BHB.Group<BHE.Node>();
                    group.Name = groupName;
                    ptForce.Objects = (BHB.Group<BHE.Node>)group;
                    ptForce.Force = force * sFac;

                    if (moment != null)
                        ptForce.Moment = moment * sFac;
                    load = ptForce;
                    break;

                case "Self-Weight":
                case "DeadLoad":
                    BHL.GravityLoad gravLoad = new BHL.GravityLoad();
                    group = new BHB.Group<BHB.BHoMObject>();
                    group.Name = groupName;
                    gravLoad.Objects = (BHB.Group<BHB.BHoMObject>)group;

                    if (force != null)
                        gravLoad.GravityDirection = force;
                    else
                        gravLoad.GravityDirection *= mag;

                    load = gravLoad;
                    break;
                case "SurfaceUDL":
                    if (force == null)
                        return "Surface Load needs force vector";
                    BHL.AreaUniformalyDistributedLoad areaLoad = new BHL.AreaUniformalyDistributedLoad();

                    areaLoad.Pressure = force * sFac;

                    group = new BHB.Group<BHE.IAreaElement>();
                    group.Name = groupName;

                    areaLoad.Objects = (BHB.Group<BHE.IAreaElement>)group;

                    load = areaLoad;
                    break;
                case "BarTemperature":
                case "BarThermal":
                    if (force == null)
                        return "Temprature Load needs vector of temperature change";
                    BHL.BarTemperatureLoad tempLoad = new BHoM.Structural.Loads.BarTemperatureLoad();

                    group = new BHB.Group<BHE.Bar>();
                    group.Name = groupName;
                    tempLoad.Objects = (BHB.Group<BHE.Bar>)group;
                    tempLoad.TemperatureChange = force;
                    load = tempLoad;
                    break;
                case "BarDilation":
                case "BarForce":
                default:
                    return "Force type not recognized or implemented yet. " +
                        "Supported Force types are BarUDL, NodeDisplacement, NodeForce, Self-Weight, SurfaceUDL, BarTemperature";
            }

            if (load == null)
                return "Unable to create load";

            load.Loadcase = loadCase;

            if (axis == "Y")
                load.Axis = BHL.LoadAxis.Local;
            else
                load.Axis = BHL.LoadAxis.Global;

            if (projected == "Y")
                load.Projected = true;
            else
                load.Projected = false;



            BHG.Project.ActiveProject.AddObject((BHB.BHoMObject)group);
            BHG.Project.ActiveProject.AddObject((BHB.BHoMObject)load);
            return load.BHoM_Guid.ToString();
        }


    }
}
