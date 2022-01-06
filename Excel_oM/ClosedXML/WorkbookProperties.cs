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

using BH.oM.Base;
using System.ComponentModel;

namespace BH.oM.Adapters.Excel
{
    [Description("Object representing the meta information of the workbook.")]
    public class WorkbookProperties : BHoMObject
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        [Description("Author of the workbook.")]
        public virtual string Author { get; set; } = "";

        [Description("Title of the workbook.")]
        public virtual string Title { get; set; } = "";

        [Description("Subject of the workbook.")]
        public virtual string Subject { get; set; } = "";

        [Description("Category of the workbook.")]
        public virtual string Category { get; set; } = "";

        [Description("Keywords related to the workbook.")]
        public virtual string Keywords { get; set; } = "";

        [Description("Comments applied to the workbook.")]
        public virtual string Comments{ get; set; } = "";

        [Description("Status of the workbook.")]
        public virtual string Status { get; set; } = "";

        [Description("Last user that modified the workbook.")]
        public virtual string LastModifiedBy { get; set; } = "";

        [Description("Company of the last user that modified the workbook.")]
        public virtual string Company { get; set; } = "";

        [Description("Manager of the last user that modified the workbook.")]
        public virtual string Manager{ get; set; } = "";

        /***************************************************/
    }
}



