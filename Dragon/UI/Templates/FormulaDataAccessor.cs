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
        // convert Guid strings to objects
        public void Store(params object[] in_)
        {
            inputs = new object[in_.Length];
            for (int i = 0; i < in_.Length; i++)
            {
                inputs[i] = Evaluate(in_[i]);
            }
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
            if (item is object[,])
            {
                // Incase T is object or something similarly cabable of
                // holding a list.
                return (T)(GetDataList<object>(index) as dynamic);
            }

            // Can't always cast directly to T from object storage type even
            // when the actual type as castable to T. So have to use `as
            // dynamic` so the cast is between the actual type of `item` to T.
            return (T)(item as dynamic);
        }

        public override List<T> GetDataList<T>(int index)
        {
            object item = inputs[index];
            if (item is List<T>)
            {
                return item as List<T>;
            }
            if (item is IEnumerable<T>)
            {
                return (item as IEnumerable<T>).ToList();
            }
            if (item is IEnumerable)
            {
                // This will flatten object[,]s
                List<T> list = new List<T>();
                foreach(object o in item as IEnumerable) {
                    if (!(o is ExcelMissing || o is ExcelEmpty))
                    {
                        list.Add((T)(o as dynamic));
                    }
                }
                return list;
            }
            return null;
        }

        public override List<List<T>> GetDataTree<T>(int index)
        {
            object item = inputs[index];
            if (item is List<List<T>>)
            {
                return item as List<List<T>>;
            }
            if (item is object[,])
            {
                // Convert 2D arrays to List<List<T>> with columns as the
                // inner list, e.g.
                //     a1 b1 c1 
                //     a2 b2 c2 
                //     a3 b3 c3 
                //       ->
                //     new List<List<T>>() {
                //         new List<T>() { a1, a2, a3 },
                //         new List<T>() { b1, b2, b3 },
                //         new List<T>() { c1, c2, c3 }
                //     }
                //
                // This is arbitrary, but it has to be one way or the other
                List<List<T>> list = new List<List<T>>();
                int height = (item as object[,]).GetLength(0);
                int width = (item as object[,]).GetLength(1);
                for(int i = 0; i < width; i++)
                {
                    list.Add(new List<T>());
                    for (int j = 0; j < height; j++)
                    {
                        object o = (item as object[,])[j, i];
                        if (!(o is ExcelMissing || o is ExcelEmpty))
                        {
                            list[i].Add((T)(o as dynamic));
                        }
                    }
                }
                return list;
            }
            if (item is IEnumerable)
            {
                return (item as IEnumerable).Cast<object>()
                    .Select(o =>
                        (o is IEnumerable) ? (o as IEnumerable)
                            .Cast<object>()
                            .Select(inner => (T)(inner as dynamic))
                            .ToList()
                            : null as List<T> )
                    .ToList();

            }
            return null;
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
                Guid id = Project.ActiveProject.IAdd(data);
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
                Guid id = Project.ActiveProject.IAdd(data);
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
                Guid id = Project.ActiveProject.IAdd(data);
                output = id.ToString();
                return true;
            } catch
            {
                output = ExcelError.ExcelErrorNA;
                return false;
            }
        }

        private object Evaluate(object input)
        {
            if (input.GetType().IsPrimitive)
            {
                return input;
            }
            if(input is string)
            {
                object obj = Project.ActiveProject.GetAny(input as string);
                return obj == null ? input : obj;
            }
            if( input is object[,])
            {
                // Keep the 2D array layout but evaluate members recursively
                // to convert Guid strings into objects from the Project
                return Evaluate(input as object[,]);
            }
            return input;
        }

        private object Evaluate(object[,] input)
        {
            int height = input.GetLength(0);
            int width = input.GetLength(1);

            object[,] evaluated = new object[height, width];
            for(int i = 0; i < width; i++)
            {
                for (int j = 0; j < height; j++)
                {
                    evaluated[j,i] = Evaluate(input[j,i]);
                }
            }
            return evaluated;
        }

    }
}