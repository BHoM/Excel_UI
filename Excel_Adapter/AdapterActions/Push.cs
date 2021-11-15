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

using BH.oM.Adapter;
using BH.oM.Base;
using System.Collections.Generic;
using System.Linq;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Public Overrides                          ****/
        /***************************************************/

        public override List<object> Push(IEnumerable<object> objects, string tag = "", PushType pushType = PushType.AdapterDefault, ActionConfig actionConfig = null)
        {
            // If unset, set the pushType to AdapterSettings' value (base AdapterSettings default is FullCRUD).
            if (pushType == PushType.AdapterDefault)
                pushType = PushType.DeleteThenCreate;

            if (pushType != PushType.DeleteThenCreate)
            {
                BH.Engine.Reflection.Compute.RecordError($"Currently Excel adapter supports only {nameof(PushType)} equal to {PushType.DeleteThenCreate}");
                return new List<object>();
            }

            IEnumerable<IBHoMObject> objectsToPush = ProcessObjectsForPush(objects, actionConfig);
            if (!objectsToPush.Any())
                new List<object>();

            bool success = ICreate(objectsToPush);

            return success ? objects.ToList() : new List<object>();
        }

        /***************************************************/
    }
}
