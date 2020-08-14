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

        [Description("Converts an Excel formula to a BHoM expression object.")]
        [Input("formula", "The formula to convert.")]
        [Output("A BHoM Expression.")]
        public static IExpression FromExcelFormula(this string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return new EmptyExpression();
            }
            if (formula[0] == '=')
            {
                return formula.Substring(1).FromExcelFormula();
            }

            int index = 0;
            return ParseAddSubtract(formula, ref index);
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        private static IExpression ParseAddSubtract(string formula, ref int index)
        {
            IExpression lhs = ParseComparitor(formula, ref index);
            while (index < formula.Length)
            {
                char c = formula[index];
                if (c == '-' || c == '+')
                {
                    string op = c.ToString();
                    index++;
                    IExpression rhs = ParseComparitor(formula, ref index);
                    lhs = new BinaryExpression
                    {
                        Operator = op,
                        Left = lhs,
                        Right = rhs
                    };
                }
                else
                {
                    return lhs;
                }
            }
            return lhs;
        }

        /*******************************************/

        private static IExpression ParseUnary(string formula, ref int index)
        {
            while (index < formula.Length && char.IsWhiteSpace(formula[index]))
            {
                index++;
            }
            char c = formula[index];
            if (c == '-' || c == '+')
            {
                string op = c.ToString();
                index++;
                IExpression rhs = ParseUnary(formula, ref index);
                return new UnaryExpression
                {
                    Operator = op,
                    Expression = rhs
                };
            }
            return ParseLeaf(formula, ref index);
        }

        private static bool IsValidChar(char c)
        {
            return "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.0123456789?_$:[]#@".Contains(c);
        }

        /*******************************************/

        private static IExpression ParseString(string formula, ref int index)
        {
            string consumed = "";
            char c;
            while (index < formula.Length)
            {
                c = formula[index];
                if(c == '"')
                {
                    if(index < formula.Length - 1 && formula[index+1] == '"')
                    {
                        consumed += c;
                        c = formula[++index];
                    }
                    else
                    {
                        index++;
                        break;
                    }
                }
                consumed += c;
                index++;
            }
            return new StringExpression { Value = consumed };
        }

        /*******************************************/

        private static IExpression ParseLeaf(string formula, ref int index)
        {
            string consumed = "";
            while (index < formula.Length && char.IsWhiteSpace(formula[index]))
            {
                index++;
            }
            
            if(index < formula.Length && formula[index] == '"')
            {
                index++;
                return ParseString(formula, ref index);
            }

            while (index < formula.Length && IsValidChar(formula[index]))
            {
                consumed += formula[index++];
            }

            // skip over trailing whitespace
            while (index < formula.Length && char.IsWhiteSpace(formula[index]))
            {
                index++;
            }

            if (formula[index] == '(')
            {
                IExpression expr;
                if (consumed.Length > 0)
                {
                    var fn = new FunctionExpression { Name = consumed };
                    do
                    {
                        index++;
                        IExpression arg = ParseAddSubtract(formula, ref index);
                        fn.Arguments.Add(arg);
                    }
                    while (formula[index] == ',');
                    expr = fn;
                }
                else
                {
                    index++;
                    expr = new ExpressionGroup
                    {
                        Expression = ParseAddSubtract(formula, ref index)
                    };
                }
                if (formula[index] != ')')
                {
                    throw new Exception("Missing closing parenthesis");
                }
                index++;
                return expr;
            }

            if (formula[index] == '{')
            {
                var expr = new ArrayExpression();
                do
                {
                    index++;
                    IExpression arg = ParseAddSubtract(formula, ref index);
                    expr.Expressions.Add(arg);
                }
                while (formula[index] == ',');
                if (formula[index] != '}')
                {
                    throw new Exception("Missing closing brace");
                }
                index++;
                return expr;
            }

            if (consumed.Length == 0)
            {
                return new EmptyExpression();
            }

            if (consumed.All(c => char.IsDigit(c) || c == '.'))
            {
                return new NumberExpression { Value = consumed };
            }

            return new ReferenceExpression { Value = consumed };
        }

        /*******************************************/

        private static IExpression ParseMultipyDivide(string formula, ref int index)
        {
            IExpression lhs = ParseUnary(formula, ref index);
            while (index < formula.Length)
            {
                char c = formula[index];
                if (c == '*' || c == '/')
                {
                    string op = c.ToString();
                    index++;
                    IExpression rhs = ParseUnary(formula, ref index);
                    lhs = new BinaryExpression
                    {
                        Operator = op,
                        Left = lhs,
                        Right = rhs
                    };
                }
                else
                {
                    return lhs;
                }
            }
            return lhs;
        }

        /*******************************************/
        
        private static IExpression ParseExponent(string formula, ref int index)
        {
            IExpression lhs = ParseMultipyDivide(formula, ref index);
            while (index < formula.Length)
            {
                char c = formula[index];
                if (c == '^')
                {
                    index++;
                    IExpression rhs = ParseMultipyDivide(formula, ref index);
                    lhs = new BinaryExpression
                    {
                        Operator = "^",
                        Left = lhs,
                        Right = rhs
                    };
                }
                else
                {
                    return lhs;
                }
            }
            return lhs;
        }

        /*******************************************/

        private static IExpression ParseConcatination(string formula, ref int index)
        {
            IExpression lhs = ParseExponent(formula, ref index);
            while (index < formula.Length)
            {
                char c = formula[index];
                if (c == '&')
                {
                    index++;
                    IExpression rhs = ParseExponent(formula, ref index);
                    lhs = new BinaryExpression
                    {
                        Operator = "&",
                        Left = lhs,
                        Right = rhs
                    };
                }
                else
                {
                    return lhs;
                }
            }
            return lhs;
        }

        /*******************************************/

        private static IExpression ParseComparitor(string formula, ref int index)
        {
            IExpression lhs = ParseConcatination(formula, ref index);
            while (index < formula.Length)
            {
                char c = formula[index];
                string op = "";
                if (c == '=' || c == '<' || c == '>')
                {
                    op += c;
                    index++;
                    if (c != '=')
                    {
                        char first = c;
                        c = formula[index];
                        if ((first == '<' && c == '>') || (c == '='))
                        {
                            op += c;
                            index++;
                        }
                    }
                    IExpression rhs = ParseConcatination(formula, ref index);
                    lhs = new BinaryExpression
                    {
                        Operator = op,
                        Left = lhs,
                        Right = rhs
                    };
                }
                else
                {
                    return lhs;
                }
            }
            return lhs;
        }

        /*******************************************/
    }
}
