/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2020, the respective contributors. All rights reserved.
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
using System.Linq.Expressions;
using System.Reflection;
using BH.Engine.Reflection;
using BH.Engine.Excel;
using System.Text.RegularExpressions;

namespace BH.UI.Excel.Templates
{
    public class FormulaDataAccessor : DataAccessor
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public FormulaDataAccessor()
        {
        }

        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/

        public override T GetDataItem<T>(int index)
        {
            Type type = typeof(T);
            object item = inputs[index];

            if (item is ExcelEmpty || item is ExcelMissing) {
                object def = defaults[index];
                return def == null ? default(T) : (T)(def as dynamic);
            }
            if (item is object[,])
            {
                // Incase T is object or something similarly cabable of
                // holding a list.
                return (T)(GetDataList<object>(index) as dynamic);
            }
            if (type.IsEnum && item is string)
            {
                return (T)Enum.Parse(type, item as string);
            }
            if (type == typeof(DateTime) && item is double)
            {
                DateTime date = DateTime.FromOADate((double)item);
                return (T)(date as dynamic);
            }
            if (type == typeof(Guid) && item is string)
            {
                return (T)(Guid.Parse(item as string) as dynamic);
            }

            // Can't always cast directly to T from object storage type even
            // when the actual type as castable to T. So have to use `as
            // dynamic` so the cast is between the actual type of `item` to T.
            return (T)(item as dynamic);
        }

        /*******************************************/

        public override List<T> GetDataList<T>(int index)
        {
            object item = inputs[index];
            if (IsBlankOrError<T>(item))
            {
                return defaults[index] as List<T>;
            }
            if (item is List<T>)
            {
                return item as List<T>;
            }
            if (item is IEnumerable<T>)
            {
                return (item as IEnumerable<T>).ToList();
            }
            if (item is IEnumerable && !(item is string))
            {
                // This will flatten object[,]s
                List<T> list = new List<T>();
                foreach (object o in item as IEnumerable)
                {
                    if (IsBlankOrError<T>(o))
                        list.Add(default(T));
                    else
                        list.Add((T)(o as dynamic));
                }
                return list;
            }
            return new List<T>() { GetDataItem<T>(index) };
        }

        /*******************************************/

        public override List<List<T>> GetDataTree<T>(int index)
        {
            object item = inputs[index];
            if (IsBlankOrError<T>(item))
            {
                return defaults[index] as List<List<T>>;
            }
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
                for (int i = 0; i < width; i++)
                {
                    list.Add(new List<T>());
                    for (int j = 0; j < height; j++)
                    {
                        object o = (item as object[,])[j, i];
                        if (IsBlankOrError<T>(o))
                            list[i].Add(default(T));
                        else
                            list[i].Add((T)(o as dynamic));
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
                            : null as List<T>)
                    .ToList();

            }
            return null;
        }

        /*******************************************/

        public static object ToExcel(object data)
        {
            try
            {
                if(data == null)
                {
                    return ExcelError.ExcelErrorNull;
                }
                if (data.GetType().IsPrimitive || data is string || data is object[,])
                {
                    return data;
                }
                if (data is Guid)
                {
                    return data.ToString();
                }
                if (data is IEnumerable && !(data is ICollection))
                {
                    return ToExcel((data as IEnumerable).Cast<object>().ToList());
                }
                if (data.GetType().IsEnum)
                {
                    return Enum.GetName(data.GetType(), data);
                }
                if (data is DateTime)
                {
                    DateTime? date = data as DateTime?;
                    if (date.HasValue)
                    {
                        return date.Value.ToOADate();
                    }
                }
                return data.GetType().ToText() + " [" + Project.ActiveProject.IAdd(data) + "]";

            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }
        }
        
        /*******************************************/

        public override bool SetDataItem<T>(int index, T data)
        {
            if (data is object[,])
            {
                output = data as object[,];
                return true;
            }
            if (output.GetLength(1) <= index)
            {
                var resized = new object[1,index + 1];
                for (int i = 0; i < output.GetLength(1); i++)
                {
                    resized[0, i] = output[0,i];
                }
                output = resized;
            }
            output[0,index] = ToExcel(data);
            return true;
        }

        /*******************************************/

        public override bool SetDataList<T>(int index, IEnumerable<T> data)
        {
            if (data is ICollection)
            {
                return SetDataItem(index, data);
            }
            return SetDataItem(index, data.ToList());
        }

        /*******************************************/

        public override bool SetDataTree<T>(int index,
            IEnumerable<IEnumerable<T>> data)
        {
            if (data is ICollection && data.All(sub => sub is ICollection))
            {
                return SetDataItem(index, data);
            }
            return SetDataItem(index, data.Select(sub => sub.ToList()).ToList());
        }

        /*******************************************/

        public void StoreDefaults(object[] params_)
        {
            // Collect default values from ParamInfo so defaultable
            // arguments can be ommited in excel
            defaults = params_;
        }

        /*******************************************/

        public virtual bool Store(string function, params object[] in_)
        {
            // Store some inputs in this DataAccessor
            // convert Guid strings to objects
            inputs = new object[in_.Length];
            for (int i = 0; i < in_.Length; i++)
            {
                inputs[i] = Evaluate(in_[i]);
            }
            ResetOutput();
            return true;
        }

        /*******************************************/

        public virtual object GetOutput()
        {
            // Retrieve the output from this DataAccessor
            var errors = Engine.Reflection.Query.CurrentEvents()
                .Where(e => e.Type == oM.Reflection.Debugging.EventType.Error);
            if (errors.Count() > 0)
            {
                string msg = errors
                    .Select(e => e.Message)
                    .Aggregate((a, b) => a + "\n" + b);
                Engine.Excel.Query.Caller().SetNote(msg);
            }
            else
            {
                Engine.Excel.Query.Caller().SetNote("");
            }

            if (output.GetLength(0) == 1 && output.GetLength(1) == 1)
            {
                return output[0, 0];
            }
            return ArrayResizer.Resize(output);
        }

        /*******************************************/

        public virtual void ResetOutput()
        {
            Engine.Excel.Query.Caller().SetNote("");
            output = new object[,] { { ExcelError.ExcelErrorNull } };
        }

        /*******************************************/

        public Tuple<Delegate, ExcelFunctionAttribute, List<object>>
            Wrap(CallerFormula caller, Expression<Action> action)
        {
            // Create a Delegate that looks like:
            //
            // (a, b, c, ...) => {
            //     accessor.ResetOutput();
            //     accessor.StoreDefaults(defaults);
            //     accessor.Store( new [] {a, b, c, ...} );
            //     action();
            //     return ToExcel(accessor.GetOutput());
            // }


            // Create an array of [n] parameters
            string fn = caller.Function;

            var rawParams = caller.Caller.InputParams;
            ParameterExpression[] lambdaParams = rawParams
                .Select(p => Expression.Parameter(typeof(object)))
                .ToArray();
            Expression newArr = Expression.NewArrayInit(
                typeof(object),
                lambdaParams
            );

            Expression defs = Expression.Constant(rawParams.Select(p => p.DefaultValue).ToArray());

            Expression accessorInstance = Expression.Constant(this);
            Type accessorType = GetType();

            // Invoke action
            Expression actionCall = Expression.Invoke(action);

            MethodInfo storeDefMethod = accessorType.GetMethod("StoreDefaults");
            Expression storeDefCall = Expression.Call(
                accessorInstance, // FormulaDataAccessor
                storeDefMethod,   // void StoreDefaults(...)
                defs
            );

            // Call FormulaDataAccessor.Store with array
            MethodInfo storeMethod = accessorType.GetMethod("Store");
            Expression storeCall = Expression.Call(
                accessorInstance, // (FormulaDataAccessor)DataAccessor
                storeMethod,      // void Store(string, object[])
                Expression.Constant(fn), // fn
                newArr            // new [] { ... }
            );

            // Return call FormulaDataAccessor.GetOutput()
            MethodInfo returnMethod = accessorType.GetMethod("GetOutput");
            Expression returnCall = Expression.Call(
                accessorInstance, // (FormulaDataAccessor)DataAccessor
                returnMethod      // object GetOutput()
            );

            // Chain them together
            Expression tree = Expression.Block(
                storeDefCall,
                Expression.Condition(
                    storeCall,
                    actionCall,
                    Expression.Empty()
                ),
                returnCall
            );
            LambdaExpression lambda = Expression.Lambda(tree, lambdaParams);

            // Compile
            var argAttrs = rawParams
                        .Select(p =>
                        {
                            string name = p.HasDefaultValue ? $"[{p.Name}]" : p.Name;
                            string postfix = string.Empty;
                            if (p.HasDefaultValue)
                            {
                                postfix += " [default: " +
                                (p.DefaultValue is string
                                    ? $"\"{p.DefaultValue}\""
                                    : p.DefaultValue == null
                                        ? "null"
                                        : p.DefaultValue.ToString()
                                ) + "]";
                            }

                            int limit = 253 - name.Length;
                            string desc = p.Description + postfix;

                            if (desc.Length >= limit) desc = p.Description.Substring(limit - postfix.Length) + postfix;

                            return new ExcelArgumentAttribute()
                            {
                                Name = name,
                                Description = desc
                            };
                        });
            if (argAttrs.Count() > 0)
            {
                string argstring = argAttrs.Select(item => item.Name).Aggregate((a, b) => $"{a}, {b}");
                if (argstring.Length >= 254)
                {
                    int i = 0;
                    argAttrs = argAttrs.Select(attr => new ExcelArgumentAttribute
                    {
                        Description = attr.Description,
                        Name = "arg" + i++
                    });
                }
            }
            return new Tuple<Delegate, ExcelFunctionAttribute, List<object>>(
                lambda.Compile(),
                GetFunctionAttribute(caller),
                argAttrs.ToList<object>()
            );
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private object Evaluate(object input)
        {
            if (input.GetType().IsPrimitive)
            {
                return input;
            }
            if (input is string)
            {
                object obj = Project.ActiveProject.GetAny(input as string);
                return obj == null ? input : obj;
            }
            if (input is object[,])
            {
                // Keep the 2D array layout but evaluate members recursively
                // to convert Guid strings into objects from the Project
                return Evaluate(input as object[,]);
            }
            return input;
        }

        /*******************************************/

        private object Evaluate(object[,] input)
        {
            int height = input.GetLength(0);
            int width = input.GetLength(1);

            object[,] evaluated = new object[height, width];
            for (int i = 0; i < width; i++)
            {
                for (int j = 0; j < height; j++)
                {
                    evaluated[j, i] = Evaluate(input[j, i]);
                }
            }
            return evaluated;
        }

        /*******************************************/

        private bool IsBlankOrError<T>(object obj)
        {
            bool isString = typeof(T) == typeof(string);

            // This will evaluate to true for "" unless T is a string
            return obj is ExcelMissing || obj is ExcelEmpty || obj is ExcelError
                || (obj is string && !isString && string.IsNullOrEmpty(obj as string));
        }

        /*******************************************/

        private ExcelFunctionAttribute GetFunctionAttribute(CallerFormula caller)
        {
            int limit = 254;
            string description = caller.Caller.Description;
            if (description.Length >= limit) description = description.Substring(0, limit-1);
            return new ExcelFunctionAttribute()
            {
                Name = caller.Function,
                Description = description,
                Category = "BHoM." + caller.Caller.Category,
                IsMacroType = true
            };
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private object[] inputs;
        private object[] defaults;
        private object[,] output = new object[,] { { null } };
    }
}

