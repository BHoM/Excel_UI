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