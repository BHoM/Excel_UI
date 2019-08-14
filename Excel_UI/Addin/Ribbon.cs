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
        private IRibbonUI m_ribbon = null;

        public override string GetCustomUI(string RibbonID)
        {
            string ribbonxml = $@"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoadRibbon' loadImage='LoadImage'>
      <ribbon>
        <tabs>
          <tab id='bhomTab' label='BHoM'>
            <group id='uninitialised' label='BHoM' getVisible='GetVisible'>
              <button id='enableBtn' label='Enable BHoM' onAction='EnableBHom' getImage='GetImage' size='large'/>
            </group>
            {AddIn.GetRibbonXml()}
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
            return ribbonxml;
        }

        public void EnableBHoM(IRibbonControl control)
        {
            AddIn.EnableBHoM((success) => {
                if(m_ribbon != null)
                {
                    m_ribbon.Invalidate();
                }
            });
        }
        
        public Bitmap GetImage(IRibbonControl control)
        {
            if (control.Id == "enableBtn") return BH.UI.Excel.Properties.Resources.BHoM_Logo;

            Templates.CallerFormula caller = AddIn.GetCaller(control.Id);
            if (caller != null) return caller.Caller.Icon_24x24;
            return null;
        }

        public bool GetVisible(IRibbonControl control)
        {
            if(control.Id == "uninitialised") return !AddIn.Enabled;
            return AddIn.Enabled;
        }

        public void OnLoadRibbon(IRibbonUI ribbon)
        {
            m_ribbon = ribbon;
        }

        public void FillFormula(IRibbonControl control)
        {
            Templates.CallerFormula caller = AddIn.GetCaller(control.Tag);
            if (caller == null) return;
            caller.Select(control.Id);
        }
    }
}
