/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2021, the respective contributors. All rights reserved.
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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using BH.Engine.Reflection;
using BH.oM.Base;
using System.Reflection;
using BH.Adapter;
using BH.oM.Data.Requests;
using BH.oM.Adapter;

namespace BH.UI.Excel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("DF6BC7B6-456C-4C5A-B591-A482F0509B18")]
    public class Adapter
    {
        /***************************************************/
        /**** Properties                                ****/
        /***************************************************/

        public string Type
        {
            get
            {
                if (m_Adapter != null)
                    return m_Adapter.GetType().ToString();
                else
                    return "null";
            }
        }


        /***************************************************/
        /**** Constructors                              ****/
        /***************************************************/

        public Adapter(string adapterName, Collection parameters = null)
        {
            Type type = BH.Engine.Reflection.Query.AdapterTypeList().Where(x => x.FullName.Contains(adapterName)).FirstOrDefault();
            if (type != null)
                m_Adapter = Helpers.RunBestComMethod(type.GetConstructors(), parameters) as BHoMAdapter;
        }

        /***************************************************/

        public Adapter(string adapterName, string filePath, Object toolkitConfig = null)
        {
            Type type = BH.Engine.Reflection.Query.AdapterTypeList().Where(x => x.FullName.Contains(adapterName)).FirstOrDefault();
            if (type != null)
                m_Adapter = BH.Engine.Adapter.Create.BHoMAdapter(type, filePath, toolkitConfig.FromCom(), true) as BHoMAdapter;
        }


        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        public Collection Push(Collection objects, string tag = "", PushType pushType = PushType.AdapterDefault, Object actionConfig = null)
        {
            if (m_Adapter == null || objects == null)
                return new Collection();

            oM.Adapter.PushType pt = oM.Adapter.PushType.AdapterDefault;
            if (!m_Adapter.SetupPushType(pushType.FromCom(), out pt))
            {
                BH.Engine.Reflection.Compute.RecordError($"Invalid `{nameof(pushType)}` input.");
                return new Collection();
            }

            ActionConfig pushConfig = null;
            if (!m_Adapter.SetupPushConfig(ComConverter.FromCom(actionConfig) as ActionConfig, out pushConfig))
            {
                BH.Engine.Reflection.Compute.RecordError($"Invalid `{nameof(actionConfig)}` input.");
                return new Collection();
            }

            List<object> result = m_Adapter.Push(objects.FromCom(), tag, pt, pushConfig);
            return result.ToCom();
        }

        /***************************************************/

        public Collection Pull(Object request = null, Object actionConfig = null)
        {
            if (m_Adapter == null)
                return new Collection();

            IRequest actualRequest = null;
            if (!m_Adapter.SetupPullRequest(ComConverter.FromCom(request) as IRequest, out actualRequest))
            {
                BH.Engine.Reflection.Compute.RecordError($"Invalid `{nameof(request)}` input.");
                return new Collection();
            }

            ActionConfig pullConfig = null;
            if (!m_Adapter.SetupPullConfig(ComConverter.FromCom(actionConfig) as ActionConfig, out pullConfig))
            {
                BH.Engine.Reflection.Compute.RecordError($"Invalid `{nameof(actionConfig)}` input.");
                return new Collection();
            }

            List<object> result = m_Adapter.Pull(actualRequest, PullType.AdapterDefault, pullConfig).ToList();
            return result.ToCom();
        }

        /***************************************************/

        public int Remove(Object request, Object actionConfig = null)
        {
            IRequest actualRequest = null;
            if (!m_Adapter.SetupRemoveRequest(ComConverter.FromCom(request) as IRequest, out actualRequest))
            {
                BH.Engine.Reflection.Compute.RecordError($"Invalid `{nameof(request)}` input.");
                return 0;
            }

            ActionConfig removeConfig = null;
            if (!m_Adapter.SetupRemoveConfig(ComConverter.FromCom(actionConfig) as ActionConfig, out removeConfig))
            {
                BH.Engine.Reflection.Compute.RecordError($"Invalid `{nameof(actionConfig)}` input.");
                return 0;
            }

            return m_Adapter.Remove(actualRequest, removeConfig);
        }


        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        protected BHoMAdapter m_Adapter = null;

        /***************************************************/
    }

}
