using System;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
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
        public static Expression<Func<Guid, T>> GuidToObject<T>() 
        {
            return x => x == Guid.Empty ? default(T) : (T)Project.ActiveProject.GetAny(x);
        }

        /*****************************************************************/

        public static Expression<Func<T, Guid>> ObjectToGuid<T>()
        {
            return x => Project.ActiveProject.IAdd(x);
        }

        ///*****************************************************************/

        public static Expression<Func<object[], List<T>>> ArrayToObjectList<T>()
        {
            return x => x.Select(y => (T)Project.ActiveProject.GetAny(y as string)).ToList();
        }

        /*****************************************************************/


        public static Expression<Func<List<T>, Guid>> ListToGuid<T>()
        {
            return x => Project.ActiveProject.Add(new ExcelList<T>() { Data = x });
        }

        /*****************************************************************/

        public static Expression<Func<IEnumerable<T>, Guid>> IEnumerableToGuid<T>()
        {
            return x => Project.ActiveProject.Add(new ExcelList<T>() { Data = x.ToList() });
        }


        /*****************************************************************/
        public static Expression<Func<Guid, BHoMGroup<T>>> GuidToBHoMGroup<T>() where T : IBHoMObject
        {
            return x => (BHoMGroup<T>)Project.ActiveProject.GetBHoM(x);
        }

        /*****************************************************************/

        public static Expression<Func<BHoMGroup<T>, Guid>> BHoMGroupToGuid<T>() where T : IBHoMObject
        {
            return x => Project.ActiveProject.Add(x);
        }

        /*****************************************************************/
    }
}
