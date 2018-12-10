using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Dragon
{

    /*****************************************************************/
    /******* Enums                                      **************/
    /*****************************************************************/

    public enum ErrorMessageHandling
    {
        ErrorMessage,   //Show the errormessages from thrown exception in the cells failing
        ErrorValue,     //Show the default error value "#VALUE" in cells failing. Default behaviour
        EmptyCell       //Leave cells failing empty
    }

    /*****************************************************************/
    /******* Static config class to handle debug configs    **********/
    /*****************************************************************/

    public static class DebugConfig
    {
        public const ErrorMessageHandling ErrorHandling = ErrorMessageHandling.ErrorValue;  //Determains what to show in cells where calculations fail
        public const bool ShowExcelDNALog = false;                                          //Show the excel dna dialog at startup, showing what methods have failed to initialize. Useful for debugging, but anoying for release
    }

    /*****************************************************************/
}
