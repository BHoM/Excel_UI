using BH.oM.UI;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon.UI.Templates
{
    public class FormulaParameter : IFormulaParameter
    {
        public ParamInfo ParamInfo { get; protected set; }
        public ExcelArgumentAttribute ArgumentAttribute
        {
            get
            {
                var attr = new ExcelArgumentAttribute();
                try
                {
                    attr.Name = ParamInfo.Name;
                    attr.Description = ParamInfo.Description;
                    if (ParamInfo.HasDefaultValue)
                    {
                        // So InteliSense makes it clear to the user that the
                        // parameter is optional.
                        attr.Name = $"[{attr.Name}]";
                        object default_ = ParamInfo.DefaultValue;
                        string def = $"[default: {default_.ToString()}]";
                        if (string.IsNullOrWhiteSpace(attr.Description))
                        {
                            attr.Description = def;
                        }
                        else
                        {
                            attr.Description += "\n" + def;
                        }
                    }
                }
                catch { }
                return attr;
            }
        }

        public FormulaParameter(ParamInfo info)
        {
            ParamInfo = info;
        }
    }
}
