using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using BH.Engine.Reflection;
using ExcelDna.Registration;
using BH.oM.Base;
using BH.oM.Geometry;
using System.Linq.Expressions;


namespace BH.UI.Dragon
{
    public class AddIn : IExcelAddIn
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

            List<string> toLoad = new List<string>() { "BHoM_Engine", "Structure_Engine", "Geometry_Engine", "Aucoustics_Engine" };

            List<MethodInfo> list = Query.BHoMMethodList().Where(x => x.IsStatic).Where(x => !x.IsGenericMethod).Where(x => toLoad.Contains(x.DeclaringType.Assembly.GetName().Name)).ToList();



            Registrations(list)
                .ProcessMapArrayFunctions(conversionConfig)
                .ProcessParameterConversions(conversionConfig)
                .RegisterFunctions();
        }

        /*****************************************************************/

        private static List<ExcelFunctionRegistration> Registrations(List<MethodInfo> methods)
        {
            List<ExcelFunctionRegistration> regs = new List<ExcelFunctionRegistration>();
            foreach (var group in methods.GroupBy(x => x.Name))
            {
                if (group.Count() == 1)
                {
                    regs.Add(ExcelFunctionRegistration(group.First(), "BH_" + group.First().Name));
                }
                else
                {
                    foreach (MethodInfo methodInfo in group)
                    {
                        string paramNames = "";
                        foreach (string s in methodInfo.GetParameters().Select(x => x.ParameterType.Name))
                        {
                            paramNames += "_" + s;
                        }
                        regs.Add(ExcelFunctionRegistration(methodInfo, "BH_" + methodInfo.Name + paramNames));
                    }
                }
            }
            return regs;
        }
        /*****************************************************************/
        private static ExcelFunctionRegistration ExcelFunctionRegistration(MethodInfo methodInfo, string name)
        {
            var paramExprs = methodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            LambdaExpression lambda = Expression.Lambda(Expression.Call(methodInfo, paramExprs), name, paramExprs);

            var allMethodAttributes = methodInfo.GetCustomAttributes(true);

            ExcelFunctionAttribute functionAttribute = null;

            foreach (var att in allMethodAttributes)
            {
                var funcAtt = att as ExcelFunctionAttribute;
                if (funcAtt != null)
                {
                    functionAttribute = funcAtt;
                    // At least ensure that name is set - from the method if need be.
                    if (string.IsNullOrEmpty(functionAttribute.Name))
                        functionAttribute.Name = name;
                }
            }
            // Check that ExcelFunctionAttribute has been set
            if (functionAttribute == null)
            {
                functionAttribute = new ExcelFunctionAttribute { Name = name };
            }

            IEnumerable<ExcelParameterRegistration> parameterRegistrations = methodInfo.GetParameters().Select(pi => new ExcelParameterRegistration(pi)).ToList();

            return new ExcelFunctionRegistration(lambda, functionAttribute, parameterRegistrations);
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



            var paramConversionConfig = new ParameterConversionConfiguration();
            // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
            paramConversionConfig.AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true))
            //Register IObject and IBHoMGeometry
            .AddParameterConversion((Guid value) => Project.ActiveProject.GetObject(value))
            .AddParameterConversion((Guid value) => Project.ActiveProject.GetGeometry(value));

            //Register type coversions for all bhom objects from guid to BHoMObjects
            foreach (Type type in ReflectionExtra.BHoMTypeList())
            {
                if (typeof(IObject).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddParameterConversion(GetParamConvGuidToBHoM(type), type);
                    var listParam = GetParamConvArrayToObjectList(type);
                    paramConversionConfig.AddParameterConversion(listParam, typeof(List<>).MakeGenericType(new Type[] { type }));
                    paramConversionConfig.AddParameterConversion(listParam, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));
                }
                else if (typeof(IBHoMGeometry).IsAssignableFrom(type))
                {
                    paramConversionConfig.AddParameterConversion(GetParamConvGuidToGeom(type), type);
                    var listParam = GetParamConvArrayToGeomList(type);
                    paramConversionConfig.AddParameterConversion(listParam, typeof(List<>).MakeGenericType(new Type[] { type }));
                    paramConversionConfig.AddParameterConversion(listParam, typeof(IEnumerable<>).MakeGenericType(new Type[] { type }));
                }
            }



            paramConversionConfig

                // Register some type conversions (note the ordering discussed above)        
                //
                //.AddParameterConversion((Guid value) => (BHoMObject)Project.ActiveProject.GetObject(value))
                //.AddParameterConversion((Guid value) => (CustomObject)Project.ActiveProject.GetObject(value))
                .AddParameterConversion((Guid value) => ((ExcelList<string>)Project.ActiveProject.GetObject(value)).Data)
                .AddParameterConversion((Guid value) => ((ExcelList<object>)Project.ActiveProject.GetObject(value)).Data)
                .AddParameterConversion((Guid value) => ((ExcelList<int>)Project.ActiveProject.GetObject(value)).Data)
                .AddParameterConversion((string value) => Guid.Parse(value));

            // This is a conversion applied to the return value of the function


            //Register type coversions for all IObjects from guid to BHoMObjects
            paramConversionConfig.AddReturnConversion((IObject value) => Project.ActiveProject.AddObject(value), true)
                .AddReturnConversion((IBHoMGeometry value) => Project.ActiveProject.AddGeometry(value), true);

            //Register type coversions for all bhom objects from guid to BHoMObjects
            foreach (Type type in ReflectionExtra.BHoMTypeList())
            {
                if (typeof(IObject).IsAssignableFrom(type))
                    paramConversionConfig.AddReturnConversion(GetParamConvBHoMToGuid(type), type);
                else if (typeof(IBHoMGeometry).IsAssignableFrom(type))
                    paramConversionConfig.AddReturnConversion(GetParamConvGeomToGuid(type), type);
            }


            paramConversionConfig
                .AddReturnConversion((List<string> value) => Project.ActiveProject.AddObject(new ExcelList<string>() { Data = value }))
                .AddReturnConversion((List<int> value) => Project.ActiveProject.AddObject(new ExcelList<int>() { Data = value }))
                .AddReturnConversion((List<object> value) => Project.ActiveProject.AddObject(new ExcelList<object>() { Data = value }))
                .AddReturnConversion((Guid value) => value.ToString())

                //  .AddParameterConversion((string value) => convert2(convert1(value)));

                // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
                // It uses the TypeConversion utility class defined in ExcelDna.Registration to get an object->string
                // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())

                // This is a pair of very generic conversions for Enum types
                .AddReturnConversion((Enum value) => value.ToString(), handleSubTypes: true)
                .AddParameterConversion(ParameterConversions.GetEnumStringConversion());

                //.AddParameterConversion((object[] input) => new Complex(TypeConversion.ConvertToDouble(input[0]), TypeConversion.ConvertToDouble(input[1])))
                //.AddNullableConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true);

            return paramConversionConfig;
        }

        /*****************************************************************/
        /******* Converter expressions                      **************/
        /*****************************************************************/

        public static Expression<Func<Guid, T>> GuidToBHom<T>() where T : IObject
        {
            return x => (T)Project.ActiveProject.GetObject(x);
        }

        /*****************************************************************/

        public static Expression<Func<T, Guid>> BHoMToGuid<T>() where T : IObject
        {
            return x => Project.ActiveProject.AddObject(x);
        }

        /*****************************************************************/

        public static Expression<Func<Guid, T>> GuidToGeom<T>() where T : IBHoMGeometry
        {
            return x => (T)Project.ActiveProject.GetGeometry(x);
        }

        /*****************************************************************/

        public static Expression<Func<T, Guid>> GeomToGuid<T>() where T : IBHoMGeometry
        {
            return x => Project.ActiveProject.AddGeometry(x);
        }

        /*****************************************************************/

        public static Expression<Func<object[], List<T>>> ArrayToObjectList<T>() where T : IObject
        {
            return x => x.Select(y => (T)Project.ActiveProject.GetObject(y as string)).ToList();
        }

        /*****************************************************************/

        public static Expression<Func<object[], List<T>>> ArrayToGeometryList<T>() where T : IBHoMGeometry
        {
            return x => x.Select(y => (T)Project.ActiveProject.GetGeometry(y as string)).ToList();
        }

        /*****************************************************************/
        /******* Converter functions                        **************/
        /*****************************************************************/

        static Func<Type, ExcelParameterRegistration, LambdaExpression> GetParamConvGuidToBHoM(Type type)
        {
            MethodInfo method = MethodGuidToBHom;
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/

        static Func<Type, ExcelReturnRegistration, LambdaExpression> GetParamConvBHoMToGuid(Type type)
        {
            MethodInfo method = MethodBHoMToGuid;
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedReturnType, unusedAttributes) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/

        static Func<Type, ExcelParameterRegistration, LambdaExpression> GetParamConvGuidToGeom(Type type)
        {
            MethodInfo method = MethodGuidToGeom;
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/

        static Func<Type, ExcelReturnRegistration, LambdaExpression> GetParamConvGeomToGuid(Type type)
        {
            MethodInfo method = MethodGeomToGuid;
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedReturnType, unusedAttributes) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/

        static Func<Type, ExcelParameterRegistration, LambdaExpression> GetParamConvArrayToObjectList(Type type)
        {
            MethodInfo method = MethodArrayToObjectList;
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/
        static Func<Type, ExcelParameterRegistration, LambdaExpression> GetParamConvArrayToGeomList(Type type)
        {
            MethodInfo method = MethodArrayToGeomList;
            MethodInfo generic = method.MakeGenericMethod(new Type[] { type });
            return (unusedParamType, unusedParamReg) => (LambdaExpression)generic.Invoke(null, new object[] { });
        }

        /*****************************************************************/
        /******* Converter Method info                      **************/
        /*****************************************************************/

        private static MethodInfo MethodGuidToBHom
        {
            get
            {
                return typeof(AddIn)
                        .GetMethods()
                        .Single(m => m.Name == "GuidToBHom" && m.IsGenericMethodDefinition && m.GetParameters().Count() == 0);
            }
        }

        /*****************************************************************/

        private static MethodInfo MethodBHoMToGuid
        {
            get
            {
                return typeof(AddIn)
                        .GetMethods()
                        .Single(m => m.Name == "BHoMToGuid" && m.IsGenericMethodDefinition && m.GetParameters().Count() == 0);
            }
        }

        /*****************************************************************/

        private static MethodInfo MethodGuidToGeom
        {
            get
            {
                return typeof(AddIn)
                        .GetMethods()
                        .Single(m => m.Name == "GuidToGeom" && m.IsGenericMethodDefinition && m.GetParameters().Count() == 0);
            }
        }

        /*****************************************************************/

        private static MethodInfo MethodGeomToGuid
        {
            get
            {
                return typeof(AddIn)
                        .GetMethods()
                        .Single(m => m.Name == "GeomToGuid" && m.IsGenericMethodDefinition && m.GetParameters().Count() == 0);
            }
        }

        /*****************************************************************/

        private static MethodInfo MethodArrayToObjectList
        {
            get
            {
                return typeof(AddIn)
                        .GetMethods()
                        .Single(m => m.Name == "ArrayToObjectList" && m.IsGenericMethodDefinition && m.GetParameters().Count() == 0);
            }
        }

        /*****************************************************************/

        private static MethodInfo MethodArrayToGeomList
        {
            get
            {
                return typeof(AddIn)
                        .GetMethods()
                        .Single(m => m.Name == "ArrayToGeometryList" && m.IsGenericMethodDefinition && m.GetParameters().Count() == 0);
            }
        }

        /*****************************************************************/


        public void AutoClose()
        {
        }

        /*****************************************************************/
    }
}
