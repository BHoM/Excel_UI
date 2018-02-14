using System;
using System.Collections.Generic;
using BH.oM.Base;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon
{
    public class ExcelDictionary<TKey,TValue> : IExcelObject
    {

        /*****************************************************************/
        /******* Public properties                          **************/
        /*****************************************************************/

        public Dictionary<TKey,TValue> Data { get; set; } = new Dictionary<TKey, TValue>();

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

            foreach (KeyValuePair<TKey,TValue> kvp in Data)
            {
                props[kvp.Key.ToString()] = kvp.Value;
            }
            return props;
        }

        /*****************************************************************/
        /******* Casting                                    **************/
        /*****************************************************************/

        public static implicit operator Dictionary<TKey,TValue>(ExcelDictionary<TKey,TValue> list)
        {
            return list.Data;
        }
    }
}

