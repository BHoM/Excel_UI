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

using BH.oM.Excel;
using BH.oM.Excel.Expressions;
using BH.oM.Reflection.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel
{
    public static partial class Convert
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        [Description("Converts an BHoM expression to an Excel Formula.")]
        [Input("expression", "The expression to convert.")]
        [Output("A Formula string.")]
        public static string IToFormula(this IExpression expression)
        {
            return ToFormula(expression as dynamic);
        }

        /*******************************************/
        
        public static string ToFormula(this NumberExpression expression)
        {
            return expression.Value;
        }

        /*******************************************/

        public static string ToFormula(this ReferenceExpression expression)
        {
            return expression.Value;
        }

        /*******************************************/

        public static string ToFormula(this StringExpression expression)
        {
            return $"\"{expression.Value}\"";
        }

        /*******************************************/

        public static string ToFormula(this FunctionExpression expression)
        {
            return expression.Name + "(" + expression.Arguments.Select(e=>e.IToFormula()).Aggregate((a,b)=>$"{a},{b}") + ")";
        }

        /*******************************************/

        public static string ToFormula(this ArrayExpression expression)
        {
            return "{" + expression.Expressions.Select(e=>e.IToFormula()).Aggregate((a,b)=>$"{a},{b}") + "}";
        }

        /*******************************************/

        public static string ToFormula(this BinaryExpression expression)
        {
            return expression.Left.IToFormula() + expression.Operator + expression.Right.IToFormula();
        }

        /*******************************************/

        public static string ToFormula(this ExpressionGroup expression)
        {
            return $"({expression.Expression.IToFormula()})";
        }

        /*******************************************/

        public static string ToFormula(this UnaryExpression expression)
        {
            return $"{expression.Operator}{expression.Expression.IToFormula()}";
        }

        /*******************************************/

        public static string ToFormula(this EmptyExpression expression)
        {
            return "";
        }

        /*******************************************/
    }
}
