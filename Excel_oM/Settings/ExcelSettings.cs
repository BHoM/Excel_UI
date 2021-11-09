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
using System.Collections.Generic;
using System.ComponentModel;

namespace BH.oM.Excel.Settings
{
    public class ExcelSettings : BHoMObject
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        [Description("Names of the worksheet(s) to read or write.")]
        public virtual List<string> Worksheets { get; set; } = null;

        [Description("Range to read or write, in standard Excel format (e.g. A1:B6).")]
        public virtual string Range { get; set; } = null;

        [Description("Styling to apply to workbook and contents.")]
        public virtual Style Style { get; set; } = new Style();

        [Description("Properties to apply to workbook and contents.")]
        public virtual WorkbookProperties WorkbookProperties { get; set; } = new WorkbookProperties();

        /***************************************************/
    }
}

