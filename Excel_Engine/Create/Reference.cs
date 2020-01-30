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
using BH.oM.Excel;
using dna = ExcelDna.Integration;

namespace BH.Engine.Excel
{
    public static partial class Create
    {
        public static Reference Reference(int row, int column)
        {
            return Reference(row, row, column, column);
        }

        public static Reference Reference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId) as dna.ExcelReference;
            return Reference(rowFirst, rowLast, columnFirst, columnLast, xlref.SheetId);
        }

        public static Reference Reference(int rowFirst, int rowLast, int columnFirst, int columnLast, string sheet)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId, sheet) as dna.ExcelReference;
            return Reference(rowFirst, rowLast, columnFirst, columnLast, xlref.SheetId);
        }

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

        public static Reference Reference(List<Rectangle> rectangles)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId) as dna.ExcelReference;
            return new Reference {
                Rectangles = rectangles,
                Sheet = xlref.SheetId
            };
        }

        public static Reference Reference(List<Rectangle> rectangles, string sheet)
        {
            dna.ExcelReference xlref = dna.XlCall.Excel(dna.XlCall.xlSheetId, sheet) as dna.ExcelReference;
            return new Reference {
                Rectangles = rectangles,
                Sheet = xlref.SheetId
            };
        }

        public static Reference Reference(List<Rectangle> rectangles, IntPtr sheetId)
        {
            return new Reference {
                Rectangles = rectangles,
                Sheet = sheetId
            };
        }
    }
}