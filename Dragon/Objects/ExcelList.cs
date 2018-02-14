using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;

namespace BH.UI.Dragon
{

    //Class used to store objects in a list in a cell. Can be list of anything, thereby slightly different from BHoMGroup.
    public class ExcelList<T> : IExcelObject
    {
        /*****************************************************************/
        /******* Public properties                          **************/
        /*****************************************************************/

        public List<T> Data { get; set; } = new List<T>();

        /*****************************************************************/
        public Guid BHoM_Guid { get; set; } = Guid.NewGuid();

        /*****************************************************************/

        public object InnerObject { get { return Data; } }

        /*****************************************************************/
        /******* Public methods                          **************/
        /*****************************************************************/


        public Dictionary<string, object> PropertyDictionary()
        {
            Dictionary<string, object> props = new Dictionary<string, object>();

            for (int i = 0; i < Data.Count; i++)
            {
                props["Item" + i] = Data[i];
            }
            return props;
        }

        /*****************************************************************/
        /******* Casting                                    **************/
        /*****************************************************************/

        public static implicit operator List<T>(ExcelList<T> list)
        {
            return list.Data;
        }
    }
}
