using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Registration;

namespace BH.UI.Dragon
{
    public class DragonFunctionExecutionHandler : FunctionExecutionHandler
    {
        public override void OnEntry(FunctionExecutionArgs args)
        {

        }

        public override void OnException(FunctionExecutionArgs args)
        {
            //Choose what to happend if an exception is thrown. Controlled from the Config class
            switch (Config.ErrorHandling)
            {
                case ErrorMessageHandling.ErrorMessage:
                    //Display the error message in the excel cell
                    args.ReturnValue = args.Exception.Message;
                    args.FlowBehavior = FlowBehavior.Return;
                    break;
                case ErrorMessageHandling.EmptyCell:
                    //Leave the cell empty
                    args.ReturnValue = "";
                    args.FlowBehavior = FlowBehavior.Return;
                    break;
                case ErrorMessageHandling.ErrorValue:
                default:
                    //Default behaviour. Results in the cell taking the value "#VALUE"
                    break;
            }


        }

        public override void OnExit(FunctionExecutionArgs args)
        {

        }

        public override void OnSuccess(FunctionExecutionArgs args)
        {

        }



        int Index;

        //Configuration part taken from the ExcelDNA.Registration examples
        // (Add a registration index just to show we can attach arbitrary data to the captured handler instance which may be created for each function.)
        // If we return the same object for every function, the object needs to be re-entrancy safe is used by IsThreadSafe functions.
        static int _index = 0;
        internal static FunctionExecutionHandler LoggingHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            return new DragonFunctionExecutionHandler { Index = _index++ };
        }
    }
}
