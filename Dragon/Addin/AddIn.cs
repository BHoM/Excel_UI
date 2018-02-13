using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.Engine.Reflection;
using ExcelDna.Registration;
using BH.oM.Base;
using BH.oM.Geometry;
using System.Linq.Expressions;
using BH.Adapter;


namespace BH.UI.Dragon
{
    public partial class AddIn : IExcelAddIn
    {
        /*****************************************************************/
        public void AutoOpen()
        {
            LoadBHomAssemblies();
            RegisterBHoMMethods();
        }

        /*****************************************************************/
        private void LoadBHomAssemblies()
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string sourceFolder = @"C:\Users\" + Environment.UserName + @"\AppData\Roaming\BHoM\BHoM_dlls";

            List<string> loadedAssemblies = ass.GetReferencedAssemblies().Select(x => x.Name + ".dll").ToList();

            //Load all BHoM dlls on opening excel
            foreach (string path in Directory.GetFiles(sourceFolder, "*.dll"))
            {
                //Check that the dll is not allready loaded
                if (loadedAssemblies.Where(x => x == Path.GetFileName(path)).Count() < 1)
                    Assembly.LoadFrom(path);
            }
        }

        /*****************************************************************/
        private void RegisterBHoMMethods()
        {
            var conversionConfig = GetParameterConversionConfig();

            //List<string> toLoad = new List<string>() { "BHoM_Engine", "Structure_Engine", "Geometry_Engine", "Aucoustics_Engine" };
            List<string> toNotLoad = new List<string>() { "Rhinoceros_Engine", "GSA_Engine", "Reflection_Engine", "Mongo_Engine", "Robot_Engine" };

            List<MethodInfo> list = Query.BHoMMethodList().Where(x => x.IsStatic).Where(x => !x.IsGenericMethod).Where(x => !toNotLoad.Contains(x.DeclaringType.Assembly.GetName().Name)).ToList();

            Type adapterType = typeof(BHoMAdapter);
            List<ConstructorInfo> adapterConstructors = Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)).SelectMany(x => x.GetConstructors()).ToList();

            IEnumerable<ExcelFunctionRegistration> registrations = Registrations(list).Concat(Registrations(adapterConstructors, "Adapter."));

            //IEnumerable<ExcelFunctionRegistration>  registrations = Registrations(typeof(Tests).GetMethods().Where(x => x.IsStatic));

            registrations
                //.ProcessMapArrayFunctions(conversionConfig)
                .ProcessParamsRegistrations()
                .ProcessParameterConversions(conversionConfig)
                .RegisterFunctions();
        }

        /*****************************************************************/

        static ParameterConversionConfiguration GetParameterConversionConfig()
        {
            // NOTE: The parameter conversion list is processed once per parameter.
            //       Parameter conversions will apply from most inside, to most outside.
            //       So to apply a conversion chain like
            //           string -> Type1 -> Type2
            //       we need to register in the (reverse) order
            //           Type1 -> Type2
            //           string -> Type1
            //
            //       (If the registration were in the order
            //           string -> Type1
            //           Type1 -> Type2
            //       the parameter (starting as Type2) would not match the first conversion,
            //       then the second conversion (Type1 -> Type2) would be applied, and no more,
            //       leaving the parameter having Type1 (and probably not eligible for Excel registration.)
            //      
            //
            //       Return conversions are also applied from most inside to most outside.
            //       So to apply a return conversion chain like
            //           Type1 -> Type2 -> string
            //       we need to register the ReturnConversions as
            //           Type1 -> Type2 
            //           Type2 -> string
            //       

            ParameterConversionConfiguration paramConversionConfig = new ParameterConversionConfiguration()
                                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true));


            //Add conversion config for valuetypes such as ints, doubles and obejcts as well as for lists of these types
            paramConversionConfig = AddStandardInputConfigurations(paramConversionConfig);

            //Method for convert functions
            IEnumerable<MethodInfo> methods = typeof(AddIn).GetMethods();

            //Add conversion to BHoM
            paramConversionConfig = AddConversionToBHom(paramConversionConfig, methods);

            //Add conversions from BHoM
            paramConversionConfig = AddConversionFromBHom(paramConversionConfig, methods);

            //Converts for lists of string,objects,doubles and bools
            paramConversionConfig = AddStandardReturnConfigurations(paramConversionConfig);

            paramConversionConfig
                // This is a pair of very generic conversions for Enum types
                .AddReturnConversion((Enum value) => value.ToString(), handleSubTypes: true)
                .AddParameterConversion(ParameterConversions.GetEnumStringConversion());

            //.AddParameterConversion((object[] input) => new Complex(TypeConversion.ConvertToDouble(input[0]), TypeConversion.ConvertToDouble(input[1])))
            //.AddNullableConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true);

            return paramConversionConfig;
        }


        /*****************************************************************/

        private static ParameterConversionConfiguration AddStandardInputConfigurations(ParameterConversionConfiguration paramConversionConfig)
        {

            // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
            paramConversionConfig.AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true));

            //Conversions for standard inputs
            paramConversionConfig.AddParameterConversion((object input) => Convert.ToInt32(input))
                .AddParameterConversion((double input) => Convert.ToInt32(input))
                .AddParameterConversion((object input) => Convert.ToDouble(input));

            //Add conversions for standard list types
            paramConversionConfig
                .AddParameterConversion((object[] inputs) => inputs.Select(x => Convert.ToInt32(x)).ToList())
                .AddParameterConversion((object[] inputs) => inputs.Select(x => Convert.ToDouble(x)).ToList())
                .AddParameterConversion((object[] inputs) => inputs.Select(x => Convert.ToInt32(x)))
                .AddParameterConversion((object[] inputs) => inputs.Select(x => Convert.ToDouble(x)))
                .AddParameterConversion((double[] inputs) => inputs.ToList())
                .AddParameterConversion((object[] inputs) => inputs.ToList())
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToList());

            return paramConversionConfig;
        }

        /*****************************************************************/

        private static ParameterConversionConfiguration AddStandardReturnConfigurations(ParameterConversionConfiguration paramConversionConfig)
        {

            //Some configurations for hwo to deal with lists of strings,ints,object and doubles
            paramConversionConfig
                .AddReturnConversion((List<string> value) => Project.ActiveProject.AddObject(new ExcelList<string>() { Data = value }))
                .AddReturnConversion((List<int> value) => Project.ActiveProject.AddObject(new ExcelList<int>() { Data = value }))
                .AddReturnConversion((List<object> value) => Project.ActiveProject.AddObject(new ExcelList<object>() { Data = value }))
                .AddReturnConversion((List<double> value) => Project.ActiveProject.AddObject(new ExcelList<double>() { Data = value }));


            return paramConversionConfig;
        }

        /*****************************************************************/

        private static ParameterConversionConfiguration AddConversionToBHom(ParameterConversionConfiguration paramConversionConfig, IEnumerable<MethodInfo> methods)
        {

            //Register IObject and IBHoMGeometry
            paramConversionConfig.AddParameterConversion((Guid value) => Project.ActiveProject.GetObject(value))
            .AddParameterConversion((Guid value) => Project.ActiveProject.GetGeometry(value));


            //MethodIfo corresponding to methods creating Expressions for the conversions
            MethodInfo guidToBhom = methods.Single(m => m.Name == "GuidToBHom");
            MethodInfo guidToGeom = methods.Single(m => m.Name == "GuidToGeom");
            MethodInfo arrToObjList = methods.Single(m => m.Name == "ArrayToObjectList");
            MethodInfo arrToGeomList = methods.Single(m => m.Name == "ArrayToGeometryList");
            MethodInfo guidToAdapter = methods.Single(m => m.Name == "GuidToAdapter");
            MethodInfo adapterToGuid = methods.Single(m => m.Name == "AdapterToGuid");
            MethodInfo guidToBHoMGroup = methods.Single(m => m.Name == "GuidToBHoMGroup");


            //Register type coversions for all bhom objects and BHoMGeometries from guid to BHoMObjects
            foreach (Type type in ReflectionExtra.BHoMTypeList())
            {
                if (typeof(IObject).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToBhom), type);
                    var listParam = GetParameterConversion(type, arrToObjList);
                    paramConversionConfig.AddParameterConversion(listParam, typeof(List<>).MakeGenericType(new Type[] { type }));
                    paramConversionConfig.AddParameterConversion(listParam, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));

                    paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToBHoMGroup), typeof(BHoMGroup<>).MakeGenericType(new Type[] { type }));

                }
                else if (typeof(IBHoMGeometry).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToGeom), type);
                    var listParam = GetParameterConversion(type, arrToGeomList);
                    paramConversionConfig.AddParameterConversion(listParam, typeof(List<>).MakeGenericType(new Type[] { type }));
                    paramConversionConfig.AddParameterConversion(listParam, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));
                }
            }


            //Add adapter converts
            Type adapterType = typeof(BHoMAdapter);

            foreach (Type type in Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)))
            {
                paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToAdapter), type);
            }


            paramConversionConfig

                // Register some type conversions for simple lists)        
                .AddParameterConversion((Guid value) => ((ExcelList<string>)Project.ActiveProject.GetObject(value)).Data)
                .AddParameterConversion((Guid value) => ((ExcelList<object>)Project.ActiveProject.GetObject(value)).Data)
                .AddParameterConversion((Guid value) => ((ExcelList<int>)Project.ActiveProject.GetObject(value)).Data)
                .AddParameterConversion((Guid value) => ((ExcelList<double>)Project.ActiveProject.GetObject(value)).Data)
                //Register type conversion from string to Guid. Needs to happend after all the bhOMtypes have been registered
                .AddParameterConversion((string value) => Guid.Parse(value));

            return paramConversionConfig;
        }

        /*****************************************************************/

        private static ParameterConversionConfiguration AddConversionFromBHom(ParameterConversionConfiguration paramConversionConfig, IEnumerable<MethodInfo> methods)
        {

            //MethodIfo corresponding to methods creating Expressions for the conversions
            MethodInfo bhomToGuid = methods.Single(m => m.Name == "BHoMToGuid");
            MethodInfo geomToGuid = methods.Single(m => m.Name == "GeomToGuid");
            MethodInfo listToGuid = methods.Single(m => m.Name == "ListToGuid");
            MethodInfo iEnumerableToGuid = methods.Single(m => m.Name == "IEnumerableToGuid");
            MethodInfo adapterToGuid = methods.Single(m => m.Name == "AdapterToGuid");
            MethodInfo bhomGroupToGuid = methods.Single(m => m.Name == "BHoMGroupToGuid");


            //Register type coversions for all IObjects from guid to BHoMObjects
            paramConversionConfig
                .AddReturnConversion((IObject value) => Project.ActiveProject.AddObject(value), true)
                .AddReturnConversion((IBHoMGeometry value) => Project.ActiveProject.AddGeometry(value), true);

            //Register type coversions for all bhom objects from guid to BHoMObjects
            foreach (Type type in ReflectionExtra.BHoMTypeList())
            {
                if (typeof(IObject).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddReturnConversion(GetReturnConversion(type, bhomToGuid), type);
                    paramConversionConfig.AddReturnConversion(GetReturnConversion(type, bhomGroupToGuid), typeof(BHoMGroup<>).MakeGenericType(new Type[] { type }));
                }
                else if (typeof(IBHoMGeometry).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddReturnConversion(GetReturnConversion(type, geomToGuid), type);
                }

                var retParamList = GetReturnConversion(type, listToGuid);

                paramConversionConfig.AddReturnConversion(retParamList, typeof(List<>).MakeGenericType(new Type[] { type }));

                var retParamIEnum = GetReturnConversion(type, iEnumerableToGuid);

                paramConversionConfig.AddReturnConversion(retParamIEnum, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));
            }

            //Register conversions for the adapters
            Type adapterType = typeof(BHoMAdapter);
            foreach (Type type in Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)))
            {
                paramConversionConfig.AddReturnConversion(GetReturnConversion(type, adapterToGuid), type);
            }

            //Add conversion from Guid to string. Needs be added after all BHoM types have been added
            paramConversionConfig.AddReturnConversion((Guid value) => value.ToString());

            return paramConversionConfig;
        }

        /*****************************************************************/
        /******* Converter functions                        **************/
        /*****************************************************************/

        static Func<Type, ExcelParameterRegistration, LambdaExpression> GetParameterConversion(Type type, MethodInfo method)
        {
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/

        static Func<Type, ExcelReturnRegistration, LambdaExpression> GetReturnConversion(Type type, MethodInfo method)
        {
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }


        /***************************************************/
        public void AutoClose()
        {
        }

        /*****************************************************************/
    }
}
