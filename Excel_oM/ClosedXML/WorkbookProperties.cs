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
using BH.oM.Base;

using System.ComponentModel;

namespace BH.oM.Excel.Settings
{
    public class WorkbookProperties : BHoMObject
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        public virtual string Author { get; set; } = "author";
        public virtual string Title { get; set; } = "title";
        public virtual string Subject { get; set; } = "subject";
        public virtual string Category { get; set; } = "category";
        public virtual string Keywords { get; set; } = "keywords";
        public virtual string Comments{ get; set; } = "comments";
        public virtual string Status { get; set; } = "status";
        public virtual string LastModifiedBy { get; set; } = "lastModified";
        public virtual string Company { get; set; } = "Buro Happold";
        public virtual string Manager{ get; set; } = "manager";

        /***************************************************/
    }
}

