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

using System;
using System.Collections.Generic;
using System.ComponentModel;
using BH.oM.Excel;
using BH.oM.Reflection.Attributes;
using dna = ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        [Description("Creates a reference object to a cell in this sheet given the row and column index.")]
        [Input("row", "The zero-indexed row number.")]
        [Input("column", "The zero-indexed column number.")]
        [Output("A Reference to the cell at those coordinates in the current sheet.")]
        public static Reference Reference(int row, int column)
        {
            return Reference(row, row, column, column);
        }

        /*******************************************/

        [Description("Creates a reference object to the rectangular range in this sheet given the row and column indeces of its corners.")]
        [Input("rowFirst", "The zero-indexed row number of the first corner.")]
        [Input("columnFirst", "The zero-indexed column number of the first corner.")]
        [Input("rowLast", "The zero-indexed row number of the last corner.")]
        [Input("columnLast", "The zero-indexed column number of the last corner.")]
        [Output("A Reference to the range at those coordinates in the current sheet.")]
        public static Reference Reference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId) as dna.ExcelReference;
            return Reference(rowFirst, rowLast, columnFirst, columnLast, xlref.SheetId);
        }

        /*******************************************/

        [Description("Creates a reference object to the rectangular range given the row and column indeces of its corners and the sheet.")]
        [Input("rowFirst", "The zero-indexed row number of the first corner.")]
        [Input("columnFirst", "The zero-indexed column number of the first corner.")]
        [Input("rowLast", "The zero-indexed row number of the last corner.")]
        [Input("columnLast", "The zero-indexed column number of the last corner.")]
        [Input("sheet", "The name of the sheet to create a range for.")]
        [Output("A Reference to the range at those coordinates.")]
        public static Reference Reference(int rowFirst, int rowLast, int columnFirst, int columnLast, string sheet)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId, sheet) as dna.ExcelReference;
            return Reference(rowFirst, rowLast, columnFirst, columnLast, xlref.SheetId);
        }

        /*******************************************/

        [Description("Creates a reference object to the rectangular range given the row and column indeces of its corners and the sheet.")]
        [Input("rowFirst", "The zero-indexed row number of the first corner.")]
        [Input("columnFirst", "The zero-indexed column number of the first corner.")]
        [Input("rowLast", "The zero-indexed row number of the last corner.")]
        [Input("columnLast", "The zero-indexed column number of the last corner.")]
        [Input("sheetId", "The ID of the sheet to create a range for.")]
        [Output("A Reference to the range at those coordinates.")]
        public static Reference Reference(int rowFirst, int rowLast, int columnFirst, int columnLast, IntPtr sheetId)
        {
            return Reference(
                 new List<Rectangle> {
                    new Rectangle {
                        RowFirst = rowFirst,
                        RowLast = rowLast,
                        ColumnFirst = columnFirst,
                        ColumnLast = columnLast
                    }
                },
                sheetId
            );
        }

        /*******************************************/

        [Description("Creates a reference object to the multi-rectangular range in this sheet given a list of Rectangles.")]
        [Input("rectangles", "A list of Rectangles.")]
        [Output("A Reference to the complex range encompased by the rectangles.")]
        public static Reference Reference(List<Rectangle> rectangles)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId) as dna.ExcelReference;
            return new Reference {
                Rectangles = rectangles,
                Sheet = xlref.SheetId
            };
        }

        /*******************************************/

        [Description("Creates a reference object to the multi-rectangular range given a list of Rectangles and the sheet.")]
        [Input("rectangles", "A list of Rectangles.")]
        [Input("sheet", "The name of the sheet to create a range for.")]
        [Output("A Reference to the complex range encompased by the rectangles.")]
        public static Reference Reference(List<Rectangle> rectangles, string sheet)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId, sheet) as dna.ExcelReference;
            return new Reference {
                Rectangles = rectangles,
                Sheet = xlref.SheetId
            };
        }

        /*******************************************/

        [Description("Creates a reference object to the multi-rectangular range given a list of Rectangles and the sheet.")]
        [Input("rectangles", "A list of Rectangles.")]
        [Input("sheetId", "The ID of the sheet to create a range for.")]
        [Output("A Reference to the complex range encompased by the rectangles.")]
        public static Reference Reference(List<Rectangle> rectangles, IntPtr sheetId)
        {
            return new Reference {
                Rectangles = rectangles,
                Sheet = sheetId
            };
        }

        /*******************************************/
    }
}