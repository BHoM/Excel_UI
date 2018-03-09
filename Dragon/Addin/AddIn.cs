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
using ExcelDna.Integration;


namespace BH.UI.Dragon
{
    public partial class AddIn : IExcelAddIn
    {
        /*****************************************************************/
        /******* Public methods                             **************/
        /*****************************************************************/
        public void AutoOpen()
        {
            LoadBHomAssemblies();

            RegisterDragonMethods();
            RegisterBHoMMethods();
            
            //Hide error box showing methods not working properly
            if(!DebugConfig.ShowExcelDNALog)
                ExcelDna.Logging.LogDisplay.Hide();
        }

        /***************************************************/

        public void AutoClose()
        {
            
        }

        /*****************************************************************/
        /******* Private methods                            **************/
        /*****************************************************************/
        private void LoadBHomAssemblies()
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string sourceFolder = @"C:\Users\" + Environment.UserName + @"\AppData\Roaming\BHoM\Assemblies";

            List<string> loadedAssemblies = ass.GetReferencedAssemblies().Select(x => x.Name + ".dll").ToList();

            //Load all BHoM dlls on opening excel
            foreach (string path in Directory.GetFiles(sourceFolder, "*.dll"))
            {
                try
                {
                    //Check that the dll is not allready loaded
                    if (loadedAssemblies.Where(x => x == Path.GetFileName(path)).Count() < 1)
                        Assembly.LoadFrom(path);
                }
                catch
                {
                    
                }
            }
        }

        /*****************************************************************/
        private void RegisterDragonMethods()
        {
            //Get out all the methods marked with the excel attributes
            IEnumerable<MethodInfo> allDragonMethods = ExcelIntegration.GetExportedAssemblies()
                .SelectMany(x => x.GetTypes().SelectMany(y => y.GetMethods(BindingFlags.Public | BindingFlags.Static)))
                .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() != null);

            List<MethodInfo> adapterMethods = new List<MethodInfo>();
            List<MethodInfo> otherMethods = new List<MethodInfo>();

            Type adapterType = typeof(Dragon.Adapter.Adapter);

            foreach (MethodInfo mi in allDragonMethods)
            {
                if (mi.DeclaringType == adapterType)
                    adapterMethods.Add(mi);
                else
                    otherMethods.Add(mi);
            }

            adapterMethods.Registrations("Adapter.", false).RegisterFunctions();
            otherMethods.Registrations("Dragon.", false).RegisterFunctions();
        }


        /*****************************************************************/
        private void RegisterBHoMMethods()
        {
            var conversionConfig = GetParameterConversionConfig();
            var functionHandlerConfig = GetFunctionExecutionHandlerConfig();

            //List<string> toLoad = new List<string>() { "BHoM_Engine", "Structure_Engine", "Geometry_Engine", "Aucoustics_Engine" };
            List<string> toNotLoad = new List<string>() { "Rhinoceros_Engine", "GSA_Engine", "Reflection_Engine", "Mongo_Engine", "Robot_Engine" };

            List<MethodInfo> list = Query.BHoMMethodList().Where(x => x.IsStatic).Where(x => !x.IsGenericMethod).Where(x => !toNotLoad.Contains(x.DeclaringType.Assembly.GetName().Name)).ToList();


            Type adapterType = typeof(BHoMAdapter);
            List<ConstructorInfo> adapterConstructors = Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)).SelectMany(x => x.GetConstructors()).ToList();

            IEnumerable<ExcelFunctionRegistration> registrations = list.Registrations().Concat(adapterConstructors.Registrations("Adapter."));

            registrations
                .ProcessParameterConversions(conversionConfig)
                //.ProcessParamsRegistrations()
                .ProcessFunctionExecutionHandlers(functionHandlerConfig)
                
                .RegisterFunctions();
        }

        /*****************************************************************/

        private static ParameterConversionConfiguration GetParameterConversionConfig()
        {
            //Note below taken from the excel registration samples. The order in which the convert methods are added is very sensitive. Take greate care if modifying any of the below
            //
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

            ParameterConversionConfiguration paramConversionConfig = new ParameterConversionConfiguration();
            //                    .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true));




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
                //Enum types using methods allready implemented in ExcelDNA.Registration for to and from string
                .AddReturnConversion((Enum value) => value.ToString(), handleSubTypes: true)
                .AddParameterConversion(ParameterConversions.GetEnumStringConversion());


            // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
            paramConversionConfig
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true));
                //.AddParameterConversion(ParameterConversions.GetNullableConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true));


            return paramConversionConfig;
        }


        /*****************************************************************/

        private static ParameterConversionConfiguration AddStandardInputConfigurations(ParameterConversionConfiguration paramConversionConfig)
        {

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
                .AddReturnConversion((List<string> value) => Project.ActiveProject.Add(new ExcelList<string>() { Data = value }))
                .AddReturnConversion((List<int> value) => Project.ActiveProject.Add(new ExcelList<int>() { Data = value }))
                .AddReturnConversion((List<object> value) => Project.ActiveProject.Add(new ExcelList<object>() { Data = value }))
                .AddReturnConversion((List<double> value) => Project.ActiveProject.Add(new ExcelList<double>() { Data = value }));


            return paramConversionConfig;
        }

        public static bool Test()
        {
            return true;
        }

        /*****************************************************************/

        private static ParameterConversionConfiguration AddConversionToBHom(ParameterConversionConfiguration paramConversionConfig, IEnumerable<MethodInfo> methods)
        {

            //Register IBHoMObject and IGeometry
            paramConversionConfig.AddParameterConversion((Guid value) => Project.ActiveProject.GetBHoM(value))
            .AddParameterConversion((Guid value) => Project.ActiveProject.GetGeometry(value));

            
            MethodInfo guidToObject = methods.Single(m => m.Name == "GuidToObject");
            MethodInfo arrToObjList = methods.Single(m => m.Name == "ArrayToObjectList");
            MethodInfo guidToBHoMGroup = methods.Single(m => m.Name == "GuidToBHoMGroup");


            //Register type coversions for all bhom objects and BHoMGeometries from guid to BHoMObjects
            foreach (Type type in ReflectionExtra.BHoMTypeList())
            {
                //TODO: switch all these type checks to only check for the empty interface once it has been implemented
                if (typeof(IBHoMObject).IsAssignableFrom(type) || typeof(IGeometry).IsAssignableFrom(type) || typeof(BH.oM.DataManipulation.Queries.IQuery).IsAssignableFrom(type) || typeof(BH.oM.Common.IResult).IsAssignableFrom(type))
                {
                    //Conversion for single objects
                    paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToObject), type);

                    //COnversion for lists and IEnumerables
                    var listParam = GetParameterConversion(type, arrToObjList);
                    paramConversionConfig.AddParameterConversion(listParam, typeof(List<>).MakeGenericType(new Type[] { type }));
                    paramConversionConfig.AddParameterConversion(listParam, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));

                    if (typeof(IBHoMObject).IsAssignableFrom(type))
                    {
                        //Conversion for BHoMGroup
                        paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToBHoMGroup), typeof(BHoMGroup<>).MakeGenericType(new Type[] { type }));
                    }
                }
            }


            //Add adapter converts
            Type adapterType = typeof(BHoMAdapter);
            foreach (Type type in Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)))
            {
                paramConversionConfig.AddParameterConversion(GetParameterConversion(type, guidToObject), type);
            }


            paramConversionConfig

                // Register some type conversions for simple lists)        
                .AddParameterConversion((Guid value) => (value == Guid.Empty ? null :(ExcelList<string>)Project.ActiveProject.GetBHoM(value)).Data)
                .AddParameterConversion((Guid value) => (value == Guid.Empty ? null : (ExcelList<object>)Project.ActiveProject.GetBHoM(value)).Data)
                .AddParameterConversion((Guid value) => (value == Guid.Empty ? null : (ExcelList<int>)Project.ActiveProject.GetBHoM(value)).Data)
                .AddParameterConversion((Guid value) => (value == Guid.Empty ? null : (ExcelList<double>)Project.ActiveProject.GetBHoM(value)).Data)
                //Register type conversion from string to Guid. Needs to happend after all the bhOMtypes have been registered
                .AddParameterConversion((string value) => value == null ? Guid.Empty : Guid.Parse(value));

            return paramConversionConfig;
        }

        /*****************************************************************/

        private static ParameterConversionConfiguration AddConversionFromBHom(ParameterConversionConfiguration paramConversionConfig, IEnumerable<MethodInfo> methods)
        {

            //MethodIfo corresponding to methods creating Expressions for the conversions
            MethodInfo objectToGuid = methods.Single(m => m.Name == "ObjectToGuid");
            MethodInfo listToGuid = methods.Single(m => m.Name == "ListToGuid");
            MethodInfo iEnumerableToGuid = methods.Single(m => m.Name == "IEnumerableToGuid");
            MethodInfo bhomGroupToGuid = methods.Single(m => m.Name == "BHoMGroupToGuid");


            //Register type coversions for all IBHoMObjects from guid to BHoMObjects
            paramConversionConfig
                .AddReturnConversion((IBHoMObject value) => Project.ActiveProject.Add(value), true)
                .AddReturnConversion((IGeometry value) => Project.ActiveProject.Add(value), true);

            //Register type coversions for all bhom objects from guid to BHoMObjects
            foreach (Type type in ReflectionExtra.BHoMTypeList())
            {
                //Conversion for single object
                paramConversionConfig.AddReturnConversion(GetReturnConversion(type, objectToGuid), type);

                //Conversion for List<T>
                var retParamList = GetReturnConversion(type, listToGuid);
                paramConversionConfig.AddReturnConversion(retParamList, typeof(List<>).MakeGenericType(new Type[] { type }));

                //Conversion for IEnumerable<T>
                var retParamIEnum = GetReturnConversion(type, iEnumerableToGuid);
                paramConversionConfig.AddReturnConversion(retParamIEnum, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));

                //COnversion for BHoMGroup<T>
                if (typeof(IBHoMObject).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddReturnConversion(GetReturnConversion(type, bhomGroupToGuid), typeof(BHoMGroup<>).MakeGenericType(new Type[] { type }));
                }
            }

            //Register conversions for the adapters
            Type adapterType = typeof(BHoMAdapter);
            foreach (Type type in Query.AdapterTypeList().Where(x => x.IsSubclassOf(adapterType)))
            {
                paramConversionConfig.AddReturnConversion(GetReturnConversion(type, objectToGuid), type);
            }

            //Add conversion from Guid to string. Needs be added after all BHoM types have been added
            paramConversionConfig.AddReturnConversion((Guid value) => value.ToString());

            return paramConversionConfig;
        }

        /*****************************************************************/
        /******* Converter functions                        **************/
        /*****************************************************************/

        private static Func<Type, ExcelParameterRegistration, LambdaExpression> GetParameterConversion(Type type, MethodInfo method)
        {
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/

        private static Func<Type, ExcelReturnRegistration, LambdaExpression> GetReturnConversion(Type type, MethodInfo method)
        {
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }


        /***************************************************/

        private static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig()
        {
            return new FunctionExecutionConfiguration()
                .AddFunctionExecutionHandler(DragonFunctionExecutionHandler.LoggingHandlerSelector);
        }


        /*****************************************************************/
    }
}
