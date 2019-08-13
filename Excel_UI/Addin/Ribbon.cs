using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
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
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoadRibbon' loadImage='LoadImage'>
      <ribbon>
        <tabs>
          <tab id='bhomTab' label='BHoM'>
            <group id='uninitialised' label='BHoM' getVisible='GetVisible'>
              <button id='enableBtn' label='Enable BHoM' onAction='EnableBHom'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
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

        public bool GetVisible(IRibbonControl control)
        {
            if(control.Id == "uninitialised") return !AddIn.Enabled;
            return AddIn.Enabled;
        }

        public void OnLoadRibbon(IRibbonUI ribbon)
        {
            m_ribbon = ribbon;
        }
    }
}
