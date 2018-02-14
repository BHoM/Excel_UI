using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon
{

    public enum ErrorMessageHandling
    {
        ErrorMessage,
        ErrorValue,
        EmptyCell
    }

    public static class Config
    {
        public const ErrorMessageHandling ErrorHandling = ErrorMessageHandling.ErrorValue;
        public const bool ShowExcelDNALog = true;
    }
}
