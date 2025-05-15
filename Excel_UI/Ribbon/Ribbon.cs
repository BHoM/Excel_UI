/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
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

using BH.UI.Excel.Templates;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;

namespace BH.UI.Excel.Addin
{
    [ComVisible(true)]
    public partial class Ribbon : ExcelRibbon
    {
        /*******************************************/
        /**** Override Methods                  ****/
        /*******************************************/

        public override string GetCustomUI(string RibbonID)
        {
            string ribbonxml = $@"
    <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
      <ribbon>
        <tabs>
          <tab id='bhomTab' label='BHoM'>
            {GetRibbonXml()}
            <group id='ui' label='UI'>
                <button id='expand' onAction='RunExpand' label='Expand list' imageMso='OutlineExpandAll' supertip='Take a list stored in a single cell and expand over multiple cells (one cell per item in the list).'/>
                <button id='condense' onAction='RunCondense' label='Condense cells' imageMso='CollapseAll' supertip='Take a group of cells and store their content as a list in a single cell.'/>
                <button id='internalise' onAction='Internalise' label='Internalise data' imageMso='RecordsSaveRecord' supertip='Save the value of the selected cells so that the values will be available when the file is reopened.' />
            </group>
            <group id='help' label='Help'>
                <button id='xlwiki' onAction='OpenLink' size='large' label='BHoM Excel Wiki' imageMso='Help' tag='{BH.Engine.Excel.Query.ExcelUIWiki()}' supertip='Go to the BHoM Excel plugin wiki to view help documentation relating to this plugin.' />
                <button id='mainwiki' onAction='OpenLink' label='BHoM Wiki' imageMso='Help' tag='{Engine.Base.Query.DocumentationURL()}' supertip='Go to the core BHoM wiki to view documentation relating the BHoM.' />
                <button id='bhomxyz' onAction='OpenLink' imageMso='GetExternalDataFromWeb' label='bhom.xyz' tag='{Engine.Base.Query.BHoMWebsiteURL()}' supertip='Visit the BHoM website.' />
            </group>
            <group id='quick' label='Quick Commands'>
                <button id='select' onAction='Select' size='large' label='Select BHoMObject' imageMso='ObjectsMultiSelect' supertip='Select BHoM Object on Connected Application UI.' />
                <button id='isolate' onAction='Isolate' size='large' label='Isolate BHoMObject' imageMso='MarginsAdjust' supertip='Isolate selectable BHoMObject on Connected Application UI.'/>
                <button id='directPull' onAction='Pull' size='large' label='Direct Pull From App' imageMso='MarkToDownloadMessageCopy' supertip='Quick pull elements from External Application.'/>
                <box id='adapterBox' boxStyle='horizontal'>
                    <button id='setAdapter' size='normal' label='Adapter' onAction='SetAdapter' imageMso='ManageQuickSteps' supertip='Establish connection configuration to External Application. Select Cell of Adapter Object' />
                    <editBox  id='adapterName' getText='GetAdapterName' enabled='false' supertip='Adapter in current use' />
                </box>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
            return ribbonxml;
        }

        /*******************************************/

        public override void OnBeginShutdown(ref Array custom)
        {
            AddIn addIn = AddIn.Instance;
            if (addIn != null)
                addIn.AutoClose();

            base.OnBeginShutdown(ref custom);
        }


        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public static string GetRibbonXml()
        {
            Dictionary<string, XmlElement> groups = new Dictionary<string, XmlElement>();
            Dictionary<string, Dictionary<int, XmlElement>> boxes = new Dictionary<string, Dictionary<int, XmlElement>>();
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("root");
            doc.AppendChild(root);
            foreach (CallerFormula caller in AddIn.CallerShells.Values)
            {
                try
                {
                    XmlElement group;
                    groups.TryGetValue(caller.Category, out group);
                    if (group == null)
                    {
                        group = (XmlElement)root.AppendChild(doc.CreateElement("group"));
                        group.SetAttribute("id", caller.Category);
                        group.SetAttribute("label", caller.Category);
                        groups.Add(caller.Category, group);
                        boxes.Add(caller.Category, new Dictionary<int, XmlElement>());
                    }
                    if (!boxes[caller.Category].ContainsKey(caller.Caller.GroupIndex))
                        boxes[caller.Category].Add(caller.Caller.GroupIndex, doc.CreateElement("box"));

                    XmlElement box = boxes[caller.Category][caller.Caller.GroupIndex];
                    box.SetAttribute("id", caller.Category + "-group" + caller.Caller.GroupIndex);
                    box.SetAttribute("boxStyle", "vertical");

                    XmlDocument tmp = new XmlDocument();
                    tmp.LoadXml(caller.GetRibbonXml());
                    box.AppendChild(doc.ImportNode(tmp.DocumentElement, true));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            foreach (var kvp in boxes)
            {
                List<int> ordered = kvp.Value.Keys.ToList();
                ordered.Sort();
                foreach (int i in ordered)
                {
                    groups[kvp.Key].AppendChild(kvp.Value[i]);
                    var sep = doc.CreateElement("separator");
                    sep.SetAttribute("id", $"sep-{kvp.Key}-{i}");
                    groups[kvp.Key].AppendChild(sep);
                }
                groups[kvp.Key].RemoveChild(groups[kvp.Key].LastChild);
            }
            return root.InnerXml;
        }

        /*******************************************/

        public Bitmap GetImage(IRibbonControl control)
        {
            if (control.Id == "enableBtn")
                return BH.UI.Excel.Properties.Resources.BHoM_Logo;

            Templates.CallerFormula caller = GetCaller(control.Id);
            if (caller != null)
                return caller.Caller.Icon_24x24;
            return null;
        }

        /*******************************************/

        public string GetContent(IRibbonControl control)
        {
            Templates.CallerFormula caller = GetCaller(control.Id);
            if (caller != null)
                return caller.GetInnerRibbonXml();
            return null;
        }

        /*******************************************/

        public static CallerFormula GetCaller(string caller)
        {
            if (AddIn.CallerShells.ContainsKey(caller))
            {
                return AddIn.CallerShells[caller];
            }
            return null;
        }

        /*******************************************/

        public void FillFormula(IRibbonControl control)
        {
            Templates.CallerFormula caller = GetCaller(control.Tag);
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

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        /*******************************************/

        private static IRibbonUI _ribbon;

        /*******************************************/
    }
}





