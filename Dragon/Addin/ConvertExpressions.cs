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

        public static Expression<Func<Guid, T>> GuidToAdapter<T>() where T : BHoMAdapter
        {
            return x => (T)Project.ActiveProject.GetAdapter(x);
        }

        /*****************************************************************/

        public static Expression<Func<T, Guid>> AdapterToGuid<T>() where T : BHoMAdapter
        {
            return x => Project.ActiveProject.AddAdapter(x);
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

        public static Expression<Func<List<T>, Guid>> ListToGuid<T>()
        {
            return x => Project.ActiveProject.AddObject(new ExcelList<T>() { Data = x });
        }

        /*****************************************************************/

        public static Expression<Func<IEnumerable<T>, Guid>> IEnumerableToGuid<T>()
        {
            return x => Project.ActiveProject.AddObject(new ExcelList<T>() { Data = x.ToList() });
        }

        /*****************************************************************/

        public static Expression<Func<string, T>> StringToEnum<T>(Type type)
        {
            return x => (T)Enum.Parse(type, x);
        }

        /*****************************************************************/
        public static Expression<Func<Guid, BHoMGroup<T>>> GuidToBHoMGroup<T>() where T : IObject
        {
            return x => (BHoMGroup<T>)Project.ActiveProject.GetObject(x);
        }

        /*****************************************************************/

        public static Expression<Func<BHoMGroup<T>, Guid>> BHoMGroupToGuid<T>() where T : IObject
        {
            return x => Project.ActiveProject.AddObject(x);
        }

        /*****************************************************************/
    }
}
