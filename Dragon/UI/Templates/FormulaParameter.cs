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
                attr.Name = ParamInfo.Name;
                attr.Description = ParamInfo.Description;
                string typeinfo = typeToString(ParamInfo.DataType);
                if (ParamInfo.HasDefaultValue)
                {
                    // So InteliSense makes it clear to the user that the
                    // parameter is optional.
                    attr.Name = $"[{attr.Name}]";
                    object default_ = ParamInfo.DefaultValue;
                    if (default_ == null)
                        default_ = "null";
                    else if (default_ is string)
                        default_ = $"\"{default_}\"";
                    string def = $"[default: {default_.ToString()}]";
                    typeinfo += " " + def;
                }
                if (string.IsNullOrWhiteSpace(attr.Description))
                {
                    attr.Description = typeinfo;
                }
                else
                {
                    attr.Description += ". " + typeinfo;
                }
                return attr;
            }
        }

        public FormulaParameter(ParamInfo info)
        {
            ParamInfo = info;
        }

        private static string typeToString(Type t)
        {
            if(t.IsGenericType)
            {
                return t.Name.Split('`').FirstOrDefault()
                    + "<"
                    + t.GenericTypeArguments
                        .Select(g => typeToString(g))
                        .Aggregate((a, b) => $"{a}, {b}")
                    + ">";
            }
            return t.Name;
        }
    }

}
