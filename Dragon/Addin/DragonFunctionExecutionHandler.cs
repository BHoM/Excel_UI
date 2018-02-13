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

            switch (Config.ErrorHandling)
            {
                case ErrorMessageHandling.ErrorMessage:
                    args.ReturnValue = args.Exception.Message;
                    args.FlowBehavior = FlowBehavior.Return;
                    break;
                case ErrorMessageHandling.EmptyCell:
                    args.ReturnValue = "";
                    args.FlowBehavior = FlowBehavior.Return;
                    break;
                case ErrorMessageHandling.ErrorValue:
                default:
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

        // The configuration part - maybe move somewhere else.
        // (Add a registration index just to show we can attach arbitrary data to the captured handler instance which may be created for each function.)
        // If we return the same object for every function, the object needs to be re-entrancy safe is used by IsThreadSafe functions.
        static int _index = 0;
        internal static FunctionExecutionHandler LoggingHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            return new DragonFunctionExecutionHandler { Index = _index++ };
        }
    }
}
