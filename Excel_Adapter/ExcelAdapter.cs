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

using System;
using System.ComponentModel;
using System.IO;
using BH.oM.Reflection.Attributes;
using BH.oM.Excel.Settings;

namespace BH.Adapter.ExcelAdapter
{
    public partial class ExcelAdapter : BHoMAdapter
    {
        /***************************************************/
        /**** Constructor                               ****/
        /***************************************************/
        [Description("Specify Excel file and properties for data transfer")]
        [Input("fileSettings", "Input the file settings to get the file name and directory the Excel Adapter should use")]
        [Input("excelSettings", "Input the additional Excel Settings the adapter should use. Default null")]
        [Output("adapter", "Adapter to Excel")]
        public ExcelAdapter(BH.oM.Adapter.FileSettings fileSettings = null, ExcelSettings excelSettings = null)
        {
            if (fileSettings == null)
            {
                BH.Engine.Reflection.Compute.RecordError("Please set the File Settings correctly to enable the Excel Adapter to work correctly");
                return;
            }

            if (!Path.HasExtension(fileSettings.FileName) || Path.GetExtension(fileSettings.FileName) != ".xlsx")
            {
                // TODO: add error about .xlsx, check if .xls and .csv are mkay?
                BH.Engine.Reflection.Compute.RecordError("File name must contain a file extension");
                return;
            }

            _fileSettings = fileSettings;
            _excelSettings = excelSettings;

            //AdapterIdName = "Excel_Adapter";
        }

        /***************************************************/
        /**** Public Adapter Methods overrides          ****/
        /***************************************************/

        

        /***************************************************/
        /**** Private Fields                            ****/
        /***************************************************/

        private BH.oM.Adapter.FileSettings _fileSettings { get; set; } = null;
        private ExcelSettings _excelSettings { get; set; } = null;
    }
}

