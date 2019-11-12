﻿using BH.oM.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel.Callers
{
    class CreateCustomCaller : UI.Components.CreateCustomCaller
    {
        /*******************************************/
        /**** Constructors                      ****/
        /*******************************************/

        public CreateCustomCaller() : base()
        {
            InputParams = new List<oM.UI.ParamInfo>();
            AddInput(0, "Properties", typeof(List<string>));
            AddInput(1, "Values", typeof(List<object>));
        }

        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public override object Run(object[] inputs)
        {
            CustomObject obj = new CustomObject();

            List<string> props = inputs[0] as List<string>;
            List<object> values = inputs[1] as List<object>;
            if (props.Count == values.Count)
            {
                for (int i = 0; i < props.Count; i++)
                    obj.CustomData[InputParams[i].Name] = inputs[i];
            }

            return obj;
        }

        /*******************************************/

        public override bool SetItem(object item)
        {
            SelectedItem = item;
            return true;
        }

        /*******************************************/
    }
}
