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
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace BH.UI.Excel
{
    //Method for automatic resizing of arrays. https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/Samples/ArrayResizer.dna
    public class ArrayResizer : XlCall
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static object Resize(object[] array, Action<ExcelReference> callback)
        {
            object[,] largeArr = new object[1, array.Length];

            for (int i = 0; i < array.Length; i++)
            {
                largeArr[0, i] = array[i];
            }

            return Resize(largeArr, callback);
        }

        /*******************************************/

        public static object Resize(object[] array)
        {
            return Resize(array, (t) => { });
        }

        /*******************************************/

        public static object Resize(object[,] array)
        {
            return Resize(array, (t) => { });
        }

        /*******************************************/

        // This function will run in the UDF context.
        // Needs extra protection to allow multithreaded use.
        public static object Resize(object[,] array, Action<ExcelReference> callback)
        {
            var caller = Excel(xlfCaller) as ExcelReference;
            var target = caller;
            if (caller == null)
                return array;

            int rows = array.GetLength(0);
            int columns = array.GetLength(1);

            if (rows == 0 || columns == 0)
                return array;

            if ((caller.RowLast - caller.RowFirst + 1 == rows) &&
                (caller.ColumnLast - caller.ColumnFirst + 1 == columns))
            {
                // Size is already OK - just return result
                return array;
            }

            var rowLast = caller.RowFirst + rows - 1;
            var columnLast = caller.ColumnFirst + columns - 1;

            // Check for the sheet limits
            if (rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 ||
                columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1)
            {
                // Can't resize - goes beyond the end of the sheet - just return #VALUE
                // (Can't give message here, or change cells)
                return ExcelError.ExcelErrorValue;
            }

            var t = target = new ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId);

            ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);

            // TODO: Add some kind of guard for ever-changing result?
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // Create a reference of the right size
                DoResize(t, callback); // Will trigger a recalc by writing formula
            });

            // Return what we have - to prevent flashing #N/A
            return array;
        }

        /*******************************************/
        /**** Private Methods                   ****/
        /*******************************************/

        static void DoResize(ExcelReference target, Action<ExcelReference> callback)
        {
            // Get the current state for reset later
            using (new ExcelEchoOffHelper())
            using (new ExcelCalculationManualHelper())
            {
                ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);

                // Get the formula in the first cell of the target
                string formula = (string)Excel(xlfGetCell, 41, firstCell);
                bool isFormulaArray = (bool)Excel(xlfGetCell, 49, firstCell);
                if (isFormulaArray)
                {
                    // Select the sheet and firstCell - needed because we want to use SelectSpecial.
                    using (new ExcelSelectionHelper(firstCell))
                    {
                        // Extend the selection to the whole array and clear
                        Excel(xlcSelectSpecial, 6);
                        ExcelReference oldArray = (ExcelReference)Excel(xlfSelection);

                        oldArray.SetValue(ExcelEmpty.Value);
                    }
                }
                // Get the formula and convert to R1C1 mode
                bool isR1C1Mode = (bool)Excel(xlfGetWorkspace, 4);
                string formulaR1C1 = formula;
                if (!isR1C1Mode)
                {
                    object formulaR1C1Obj;
                    XlReturn formulaR1C1Return = TryExcel(xlfFormulaConvert, out formulaR1C1Obj, formula, true, false, ExcelMissing.Value, firstCell);
                    if (formulaR1C1Return != XlReturn.XlReturnSuccess || formulaR1C1Obj is ExcelError)
                    {
                        string firstCellAddress = (string)Excel(xlfReftext, firstCell, true);
                        Excel(xlcAlert, "Cannot resize array formula at " + firstCellAddress + " - formula might be too long when converted to R1C1 format.");
                        firstCell.SetValue("'" + formula);
                        return;
                    }
                    formulaR1C1 = (string)formulaR1C1Obj;
                }
                // Must be R1C1-style references
                object ignoredResult;
                //Debug.Print("Resizing START: " + target.RowLast);
                XlReturn formulaArrayReturn = TryExcel(xlcFormulaArray, out ignoredResult, formulaR1C1, target);
                //Debug.Print("Resizing FINISH");

                // TODO: Find some dummy macro to clear the undo stack

                if (formulaArrayReturn != XlReturn.XlReturnSuccess)
                {
                    string firstCellAddress = (string)Excel(xlfReftext, firstCell, true);

                    Excel(xlcAlert, "Cannot resize array formula at " + firstCellAddress + " - result might overlap another array.");
                    // Might have failed due to array in the way.
                    firstCell.SetValue("'" + formula);
                }
                else
                {
                    callback(target);
                }
            }
        }

        /*******************************************/
    }
}
