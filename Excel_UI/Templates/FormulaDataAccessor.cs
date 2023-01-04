/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2023, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *
 *
 * The BHoM is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3.0 of the License, or
 * (at your option) any later version.
 *
 * The BHoM is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.
 */

using BH.oM.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using BH.Engine.Reflection;
using System.Reflection;
using System.ComponentModel;

namespace BH.UI.Excel.Templates
{
    public class FormulaDataAccessor : IDataAccessor
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public List<object> Outputs { get; private set; } = new List<object> { ExcelError.ExcelErrorNull };

        public List<string> OutputNames { get; set; } = new List<string>();

        public bool TransposeOutputs { get; set; } = false;


        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        public virtual void SetInputs(List<object> inputs, List<object> defaultValues)
        {
            m_Inputs = inputs.Select(x => AddIn.FromExcel(x)).ToList();
            m_Defaults = defaultValues.Select(x => AddIn.FromExcel(x)).ToList();

            AddIn.WriteNote("");
            Outputs = new List<object> { ExcelError.ExcelErrorNull };
        }

        /*******************************************/

        public virtual object GetOutputs()
        {
            // Retrieve the output from this DataAccessor
            if (Outputs.Count == 0)
                return ExcelError.ExcelErrorNull;
            if (Outputs.Count == 1)
                return AddIn.ToExcel(Outputs[0]);
            else
            {
                List<List<object>> result = new List<List<object>> { Outputs };
                if (OutputNames.Count == Outputs.Count)
                    result.Insert(0, OutputNames.ToList<object>());

                if (TransposeOutputs)
                {
                    result = result.SelectMany(row => row.Select((value, index) => new { value, index }))
                        .GroupBy(cell => cell.index, cell => cell.value)
                        .Select(g => g.ToList()).ToList();
                }

                return AddIn.ToExcel(result);
            }    
        }


        /*******************************************/
        /**** IDataAccessor Methods             ****/
        /*******************************************/

        public virtual T GetDataItem<T>(int index)
        {
            Type type = typeof(T);
            object item = m_Inputs[index];

            if (IsBlankOrError<T>(item))
            {
                object def = m_Defaults[index];
                return (def == null || def is DBNull) ? default(T) : (T)(def as dynamic);
            }
            else if (item is object[,])
                return (T)(GetDataList<object>(index) as dynamic); // Incase T is object or something similarly cabable of holding a list.
            else
                return CastItem<T>(item);
        }

        /*******************************************/

        public virtual List<T> GetDataList<T>(int index)
        {
            object item = m_Inputs[index];
            if (IsBlankOrError<T>(item))
                return m_Defaults[index] as List<T>;
            else if (item is List<T>)
                return item as List<T>;
            else if (item is IEnumerable<T>)
                return (item as IEnumerable<T>).ToList();
            else if (item is IEnumerable && !(item is string))
            {
                // This will flatten object[,]s
                List<T> list = new List<T>();
                foreach (object o in item as IEnumerable)
                {
                    if (IsBlankOrError<T>(o))
                        list.Add(default(T));
                    else
                        list.Add(CastItem<T>(o));
                }
                return list;
            }
            else
                return new List<T>() { GetDataItem<T>(index) };
        }

        /*******************************************/

        public virtual List<List<T>> GetDataTree<T>(int index)
        {
            object item = m_Inputs[index];
            if (IsBlankOrError<T>(item))
                return m_Defaults[index] as List<List<T>>;
            else if (item is List<List<T>>)
                return item as List<List<T>>;
            else if (item is object[,])
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
                for (int i = 0; i < width; i++)
                {
                    list.Add(new List<T>());
                    for (int j = 0; j < height; j++)
                    {
                        object o = (item as object[,])[j, i];
                        if (IsBlankOrError<T>(o))
                            list[i].Add(default(T));
                        else
                            list[i].Add(CastItem<T>(o));
                    }
                }
                return list;
            }
            else if (item is IEnumerable)
            {
                return (item as IEnumerable).Cast<object>()
                    .Select(o =>
                        (o is IEnumerable) ?
                            (o as IEnumerable).Cast<object>().Select(inner => CastItem<T>(inner)).ToList()
                            : null as List<T>)
                    .ToList();
            }
            else
                return null;
        }

        /*******************************************/

        public virtual List<object> GetAllData(int index)
        {
            return GetDataList<object>(index);
        }

        /*******************************************/

        public virtual bool SetDataItem<T>(int index, T data)
        {
            while (Outputs.Count <= index)
                Outputs.Add(null);

            Outputs[index] = data;
            return true;
        }

        /*******************************************/

        public virtual bool SetDataList<T>(int index, IEnumerable<T> data)
        {
            if (data is ICollection)
                return SetDataItem(index, data);
            else
                return SetDataItem(index, data.ToList());
        }

        /*******************************************/

        public virtual bool SetDataTree<T>(int index, IEnumerable<IEnumerable<T>> data)
        {
            if (data is ICollection && data.All(sub => sub is ICollection))
                return SetDataItem(index, data);
            else
                return SetDataItem(index, data.Select(sub => sub.ToList()).ToList());
        }


        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private bool IsBlankOrError<T>(object obj)
        {
            // This will evaluate to true for "" unless T is a string
            return obj is ExcelMissing || obj is ExcelEmpty || obj is ExcelError
                || (obj is string && typeof(T) != typeof(string) && string.IsNullOrEmpty(obj as string));
        }

        /*******************************************/

        private T CastItem<T>(object item)
        {
            Type type = typeof(T); 

            if (type.IsEnum && item is string)
                return Engine.Excel.Compute.ParseEnum<T>(item as string);
            else if (type == typeof(DateTime) && item is double)
            {
                DateTime date = DateTime.FromOADate((double)item);
                return (T)(date as dynamic);
            }
            else if (type == typeof(Guid) && item is string)
                return (T)(Guid.Parse(item as string) as dynamic);
            else if (type == typeof(string) && !(item is string))
                return item?.ToString() as dynamic;
            else
            {
                // Can't always cast directly to T from object storage type even
                // when the actual type as castable to T. So have to use `as
                // dynamic` so the cast is between the actual type of `item` to T.
                return (T)(item as dynamic);
            }
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private List<object> m_Inputs { get; set; } = new List<object> { };
        private List<object> m_Defaults { get; set; } = new List<object> { };

        /*******************************************/
    }
}




