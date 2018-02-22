using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BH.oM.Base;

namespace BH.UI.Dragon
{

    //Class used to store objects in a tuple
    public class ExcelTuple<T1,T2> :  IExcelObject
    {

        /*****************************************************************/
        /******* Public properties                          **************/
        /*****************************************************************/

        public Tuple<T1, T2> Data { get; set; } = null;

        /*****************************************************************/
        public Guid BHoM_Guid { get; set; } = Guid.NewGuid();

        /*****************************************************************/

        public object InnerObject { get { return Data; } }

        /*****************************************************************/
        /******* Constructors                               **************/
        /*****************************************************************/
        public ExcelTuple()
        {
        }

        /*****************************************************************/

        public ExcelTuple(T1 item1, T2 item2)
        {
            Data = new Tuple<T1, T2>(item1, item2);
        }

        /*****************************************************************/
        /******* Public methods                          **************/
        /*****************************************************************/


        public Dictionary<string, object> PropertyDictionary()
        {
            Dictionary<string, object> props = new Dictionary<string, object>();

            props["Item1"] = Data.Item1;
            props["Item2"] = Data.Item2;
            return props;
        }

        /*****************************************************************/
        /******* Casting                                    **************/
        /*****************************************************************/

        public static implicit operator Tuple<T1,T2>(ExcelTuple<T1,T2> tuple)
        {
            return tuple.Data;
        }
    }
}
