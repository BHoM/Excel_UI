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

using BH.UI.Excel.Templates;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage'>
      <ribbon>
        <tabs>
          <tab id='bhomTab' label='BHoM'>
            {GetRibbonXml()}
            <group id='ui' label='UI'>
                <button id='internalise' onAction='Internalise' label='Internalise data' imageMso='RecordsSaveRecord' supertip='Save the value of the selected cells so it will stil lbe available when the file is reopened.' />
            </group>
            <group id='help' label='Help'>
                <button id='xlwiki' onAction='OpenLink' size='large' label='BHoM Excel Wiki' imageMso='Help' tag='https://github.com/BHoM/Excel_Toolkit/wiki' supertip='Go to the BHoM Excel plugin wiki to view help documentation relating to this plugin.' />
                <button id='mainwiki' onAction='OpenLink' label='BHoM Wiki' imageMso='Help' tag='https://github.com/BHoM/documentation/wiki' supertip='Go to the core BHoM wiki to view documentation relating the BHoM.' />
                <button id='bhomxyz' onAction='OpenLink' imageMso='GetExternalDataFromWeb' label='bhom.xyz' tag='https://bhom.xyz' supertip='Visit the BHoM website.' />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
            return ribbonxml;
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
            foreach (CallerFormula caller in AddIn.Callers.Values)
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
            if (AddIn.Callers.ContainsKey(caller))
            {
                return AddIn.Callers[caller];
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
    }
}
