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
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BH.Adapter.ExcelAdapter
{
    public partial class ExcelAdapter
    {
        protected override IEnumerable<IBHoMObject> IRead(Type type, IList ids, ActionConfig actionConfig = null)
        {
            IEnumerable<BHoMObject> everything = ReadXlsx();

            if (type != null)
                everything = everything.Where(x => type.IsAssignableFrom(x.GetType()));

            if (ids != null)
            {
                HashSet<Guid> toDelete = new HashSet<Guid>(ids.Cast<Guid>());
                everything = everything.Where(x => !toDelete.Contains((Guid)x.CustomData[AdapterIdName]));
            }


            return everything;
        }


        private IEnumerable<BHoMObject> ReadXlsx()
        {
            string[] json = File.ReadAllLines(m_FilePath);
            var converted = json.Select(x => Engine.Serialiser.Convert.FromJson(x) as BHoMObject).Where(x => x != null);
            if (converted.Count() < json.Count())
                BH.Engine.Reflection.Compute.RecordWarning("Could not convert some object to BHoMObject.");
            return converted;
        }
    }
}

