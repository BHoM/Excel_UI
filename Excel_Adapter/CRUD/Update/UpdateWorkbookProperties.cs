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

using BH.Engine.Adapter;
using BH.oM.Adapter;
using BH.oM.Data.Collections;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using BH.Engine.Data;
using System.Reflection;
using System.Threading;
using BH.oM.Base;
using BH.Engine.Base;
using BH.oM.Adapters.Excel;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter
    {
        /***************************************************/
        /**** Private Methods                           ****/
        /***************************************************/

        public void Update(IXLWorkbook workbook, WorkbookProperties properties)
        {
            if (workbook != null && properties != null)
            {
                workbook.Properties.Author = properties.Author;
                workbook.Properties.Title = properties.Title;
                workbook.Properties.Subject = properties.Subject;
                workbook.Properties.Category = properties.Category;
                workbook.Properties.Keywords = properties.Keywords;
                workbook.Properties.Comments = properties.Comments;
                workbook.Properties.Status = properties.Status;
                workbook.Properties.LastModifiedBy = properties.LastModifiedBy;
                workbook.Properties.Company = properties.Company;
                workbook.Properties.Manager = properties.Manager;
            }
        }

        /***************************************************/
    }
}


