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

using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Addin
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public override string GetCustomUI(string RibbonID)
        {
            string ribbonxml = $@"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoadRibbon' loadImage='LoadImage'>
      <ribbon>
        <tabs>
          <tab id='bhomTab' label='BHoM'>
            {AddIn.GetRibbonXml()}
            <group id='help' label='Help'>
                <button id='xlwiki' onAction='OpenLink' size='large' label='BHoM Excel Wiki' imageMso='Help' tag='https://github.com/BHoM/Excel_Toolkit/wiki' supertip='Go to the BHoM Excel plugin wiki to view help documentation relating to this plugin' />
                <button id='mainwiki' onAction='OpenLink' label='BHoM Wiki' imageMso='Help' tag='https://github.com/BHoM/documentation/wiki' supertip='Go to the core BHoM wiki to view documentation relating the BHoM' />
                <button id='bhomxyz' onAction='OpenLink' imageMso='GetExternalDataFromWeb' label='bhom.xyz' tag='https://bhom.xyz' supertip='Visit the BHoM website' />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
            return ribbonxml;
        }

        /*******************************************/

        public Bitmap GetImage(IRibbonControl control)
        {
            if (control.Id == "enableBtn")
                return BH.UI.Excel.Properties.Resources.BHoM_Logo;

            Templates.CallerFormula caller = AddIn.GetCaller(control.Id);
            if (caller != null)
                return caller.Caller.Icon_24x24;
            return null;
        }

        /*******************************************/

        public string GetContent(IRibbonControl control)
        {
            Templates.CallerFormula caller = AddIn.GetCaller(control.Id);
            if (caller != null)
                return caller.GetInnerRibbonXml();
            return null;
        }

        /*******************************************/

        public void OnLoadRibbon(IRibbonUI ribbon)
        {
            m_Ribbon = ribbon;
        }

        /*******************************************/

        public void FillFormula(IRibbonControl control)
        {
            Templates.CallerFormula caller = AddIn.GetCaller(control.Tag);
            if (caller == null)
                return;
            caller.Select(control.Id);
        }

        /*******************************************/

        public void OpenLink(IRibbonControl control)
        {
            System.Diagnostics.Process.Start(control.Tag);
        }

        /*******************************************/
        /**** Private Fields                    ****/
        /*******************************************/

        private IRibbonUI m_Ribbon = null;

        /*******************************************/
    }
}
