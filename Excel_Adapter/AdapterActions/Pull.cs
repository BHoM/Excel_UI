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

using BH.Adapter;
using BH.Engine.Adapter;
using BH.oM.Adapter;
using BH.oM.Base;
using BH.oM.Data.Requests;
using System.Collections.Generic;
using System.IO;

namespace BH.Adapter.Excel
{
    public partial class ExcelAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Method Overrides                          ****/
        /***************************************************/

        public override IEnumerable<object> Pull(IRequest request, PullType pullOption = PullType.AdapterDefault, ActionConfig actionConfig = null)
        {
            if (request == null)
            {
                BH.Engine.Base.Compute.RecordError("The request has not been provided for Pull action.");
                return new List<object>();
            }

            if (!File.Exists(m_FileSettings.GetFullFileName()))
            {
                BH.Engine.Base.Compute.RecordError("No file exists under the location specified in the settings.");
                return new List<IBHoMObject>();
            }

            return Read(request);
        }

        /***************************************************/
    }
}



