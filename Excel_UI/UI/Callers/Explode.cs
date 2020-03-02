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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Callers
{
    class ExplodeCaller : BH.UI.Templates.MethodCaller
    {
        /*******************************************/
        /**** Properties                        ****/
        /*******************************************/

        public override System.Drawing.Bitmap Icon_24x24
        {
            get
            {
                return m_Native.Icon_24x24;
            }
        }

        public override string Name
        {
            get
            {
                return m_Native.Name;
            }
        }

        public override string Category
        {
            get
            {
                return m_Native.Category;
            }
        }

        public override string Description
        {
            get
            {
                return m_Native.Description;
            }
        }

        public override int GroupIndex
        {
            get
            {
                return m_Native.GroupIndex;
            }
        }

        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public ExplodeCaller() : base(typeof(Methods.Properties).GetMethod("Explode"))
        {
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private UI.Components.ExplodeCaller m_Native = new UI.Components.ExplodeCaller();

        /*******************************************/
    }
}
