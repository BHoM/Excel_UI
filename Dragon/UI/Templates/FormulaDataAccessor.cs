using BH.oM.Base;
using BH.oM.UI;
using BH.UI.Templates;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace BH.UI.Dragon.UI.Templates
{
    public class FormulaDataAccessor : DataAccessor
    {
        private object[] inputs;
        private object[] defaults;
        private object output;

        public FormulaDataAccessor(IEnumerable<ParamInfo> params_)
        {
            // Collect default values from ParamInfo so defaultable
            // arguments can be ommited in excel
            defaults = params_.Select(p => p.DefaultValue).ToArray();
        }

        // Store some inputs in this DataAccessor
        public void Store(params object[] in_)
        {
            inputs = in_;
        }

        // Retrieve the output from this DataAccessor
        public object GetOutput()
        {
            if (output == null)
            {
                return ExcelError.ExcelErrorNull;
            }
            return output;
        }
        
        public override T GetDataItem<T>(int index)
        {
            object item = inputs[index];
            if(item is ExcelEmpty || item is ExcelMissing) {
                return (T)defaults[index];
            }
            if (item is string)
            {
                var stored = (T)Project.ActiveProject.GetAny(item as string);
                if (stored != null) return stored;
            }

            // Can't always cast directly to T from object storage type even
            // when the actual type as castable to T. So have to use `as
            // dynamic` so the cast is between the actual type of `item` to T.
            return (T)(item as dynamic);
        }

        public override List<T> GetDataList<T>(int index)
        {
            object item = inputs[index];
            if (item is string)
            {
                string id = inputs[index] as string;
                object obj = Project.ActiveProject.GetAny(id);
                return (obj as IEnumerable).Cast<T>().ToList();
            } else if (item is object[,])
            {
                object[,] list = item as object[,];
                if(typeof(T).IsPrimitive) 
                {
                    return list.Cast<object>()
                        // As above
                        .Select(o=>(T)(o as dynamic))
                        .ToList();
                }
                else if (typeof(T).Equals(typeof(string)))
                {

                    return list.Cast<object>()
                        .Select(o => o.ToString())
                        .Cast<T>() // We know T == string but compiler doesn't
                        .ToList();
                }
                // Otherwise try to retrieve objects from the Project
                List<T> l = new List<T>();
                foreach ( var listitem in list )
                {
                    if( listitem is string )
                    {
                        T stored = (T)Project.ActiveProject
                            .GetAny(listitem as string);
                        if (stored != null) l.Add(stored);
                    }
                }
                return l;
            }
            return null;
        }

        public override List<List<T>> GetDataTree<T>(int index)
        {
            string id = inputs[index] as string;
            object obj = Project.ActiveProject.GetAny(id);
            return (obj as IEnumerable).Cast<List<T>>().ToList();
        }

        public override bool SetDataItem<T>(int index, T data)
        {
            try
            {
                if (data.GetType().IsPrimitive || data is string)
                {
                    output = data;
                    return true;
                }
                Guid id = Project.ActiveProject.Add(data as dynamic);
                output = id.ToString();
                return true;
            } catch
            {
                output = ExcelError.ExcelErrorNA;
                return false;
            }
        }

        public override bool SetDataList<T>(int index, IEnumerable<T> data)
        {
            try
            {
                Guid id = Project.ActiveProject.Add(data as dynamic);
                output = id.ToString();
                return true;
            } catch
            {
                output = ExcelError.ExcelErrorNA;
                return false;
            }
        }

        public override bool SetDataTree<T>(int index,
            IEnumerable<IEnumerable<T>> data)
        {
            try
            {
                Guid id = Project.ActiveProject.Add(data as dynamic);
                output = id.ToString();
                return true;
            } catch
            {
                output = ExcelError.ExcelErrorNA;
                return false;
            }
        }
    }
}
