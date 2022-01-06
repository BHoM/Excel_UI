/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2022, the respective contributors. All rights reserved.
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

using System;
using BH.oM.Base;
using BH.UI.Excel.Templates;
using BH.UI.Base;
using System.Linq;
using BH.Engine.Reflection;
using Microsoft.Office.Core;
using System.Collections.Generic;
using BH.UI.Base.Components;
using System.Linq.Expressions;
using BH.oM.UI;
using System.Reflection;
using ExcelDna.Integration;

namespace BH.UI.Excel.Components
{
    public class CreateCustomFormula : CallerFormula
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override Caller Caller { get; } = new CreateCustomCaller(false);

        public override string Function => GetName();


        /*******************************************/
        /**** Override Methods                  ****/
        /*******************************************/

        public override string GetName()
        {
            return  "Create.CustomObject";
        }

        /*******************************************/

        protected override MethodCallExpression GetMethodCall(ref List<ParamInfo> inputs, ref ParameterExpression[] lambdaParams)
        {
            int nbInputs = Caller.InputParams.Count;
            NewArrayExpression array = Expression.NewArrayInit(typeof(object), lambdaParams);

            List<ParamInfo> extraInputs = new List<ParamInfo>
            {
                new ParamInfo { DataType = typeof(object), DefaultValue = null, HasDefaultValue = true, IsRequired = false, Name = "_inputTypes", Description = "Force the input values to be of the given types." },
            };

            inputs.AddRange(extraInputs);
            lambdaParams = lambdaParams.Concat(new ParameterExpression[] { Expression.Parameter(typeof(object)) }).ToArray();
            MethodInfo method = GetType().GetMethod("RunWithForcedTypes");

            return Expression.Call(Expression.Constant(this), method, array, lambdaParams[nbInputs]);
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public object RunWithForcedTypes(object[] inputs, object forcedTypes = null)
        {
            if (m_DataAccessor != null && forcedTypes != null && inputs != null && inputs.Length >= 2 && !(forcedTypes is ExcelMissing))
            {
                m_DataAccessor.SetInputs(new List<object> { inputs[1], forcedTypes }, new List<object> { new List<object>(), new List<Type>() });
                List<object> objects = m_DataAccessor.GetDataList<object>(0);
                List<Type> types = m_DataAccessor.GetDataList<Type>(1);

                m_DataAccessor.SetInputs(objects, objects.Select(x => null as object).ToList());

                for (int i = 0; i < Math.Min(objects.Count, types.Count); i++)
                {

                    if (types[i] != null && objects[i] != null)
                    {
                        Type type = types[i];
                        if (!m_GetDataItemAccessors.ContainsKey(type))
                            m_GetDataItemAccessors[type] = GetType().GetMethod("CallGetDataItem").MakeGenericMethod(new Type[] { type })?.ToFunc();

                        Func<object[], object> getDataItem = m_GetDataItemAccessors[type];
                        if (getDataItem != null)
                        {
                            object result = getDataItem(new object[] { m_DataAccessor, i });
                            if (result != null)
                                objects[i] = result;
                        }
                    }
                }

                inputs[1] = objects;
            }

            return Run(inputs);
        }

        /*******************************************/

        public static object CallGetDataItem<T>(FormulaDataAccessor accessor, int index)
        {
            if (accessor != null)
                return accessor.GetDataItem<T>(index);
            else
                return null;
        }


        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        protected static Dictionary<Type, Func<object[], object>> m_GetDataItemAccessors = new Dictionary<Type, Func<object[], object>>();

        /*******************************************/
    }
}



