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
    // Select an ExcelReference (perhaps on another sheet) allowing changes to be made there.
    // On clean-up, resets all the selections and the active sheet.
    // Should not be used if the work you are going to do will switch sheets, amke new sheets etc.
    public class ExcelSelectionHelper : XlCall, IDisposable
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ExcelSelectionHelper(ExcelReference refToSelect)
        {
            // Remember old selection state on the active sheet
            oldSelectionOnActiveSheet = Excel(xlfSelection);
            oldActiveCellOnActiveSheet = Excel(xlfActiveCell);

            // Switch to the sheet we want to select
            string refSheet = (string)Excel(xlSheetNm, refToSelect);
            Excel(xlcWorkbookSelect, new object[] { refSheet });

            // record selection and active cell on the sheet we want to select
            oldSelectionOnRefSheet = Excel(xlfSelection);
            oldActiveCellOnRefSheet = Excel(xlfActiveCell);

            // make the selection
            Excel(xlcFormulaGoto, refToSelect);
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void Dispose()
        {
            // Reset the selection on the target sheet
            Excel(xlcSelect, oldSelectionOnRefSheet, oldActiveCellOnRefSheet);

            // Reset the sheet originally selected
            string oldActiveSheet = (string)Excel(xlSheetNm, oldSelectionOnActiveSheet);
            Excel(xlcWorkbookSelect, new object[] { oldActiveSheet });

            // Reset the selection in the active sheet (some bugs make this change sometimes too)
            Excel(xlcSelect, oldSelectionOnActiveSheet, oldActiveCellOnActiveSheet);
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        object oldSelectionOnActiveSheet;
        object oldActiveCellOnActiveSheet;

        object oldSelectionOnRefSheet;
        object oldActiveCellOnRefSheet;

        /*******************************************/
    }
}
