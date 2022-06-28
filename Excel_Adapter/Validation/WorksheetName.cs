/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2022, the respective contributors. All rights reserved.
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

using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace BH.Adapter.Excel
{
    public partial class Validation
    {
        /***************************************************/
        /**** Public Methods                            ****/
        /***************************************************/

        
        public static string WorksheetName(string name, IXLWorkbook workbook)
        {
            Dictionary<string, bool> nameChecks = new Dictionary<string, bool>();

            string workSheetName = !IsValidName(name, workbook, out nameChecks) ? ModifyWorksheetName(name, nameChecks, workbook) : name;

            if (workSheetName != name)
            {
                BH.Engine.Base.Compute.RecordError("Name of worksheet has been adjusted to a name that which is compatible with Excel naming limitations.");
            }

            return workSheetName;
        }

        private static bool IsValidName(string workSheetName, IXLWorkbook workbook, out Dictionary<string, bool> nameChecks)
        {
            nameChecks = new Dictionary<string, bool>();

            bool isUnique = IsUnique(workSheetName, workbook);
            bool isWithinCharacterLimit = IsWithinCharacterLimit(workSheetName);
            bool isNotBlank = !string.IsNullOrWhiteSpace(workSheetName);
            bool isNotReservedWord = IsNotReservedWord(workSheetName);
            bool isValidCharacters = IsValidCharacters(workSheetName);
            bool isNotBeginWithInvalidCharacter = IsNotBeginOrEndWithInvalidCharacter(workSheetName);

            if (!isUnique)
            {
                nameChecks.Add("isUnique", isUnique);
            }

            if (!isWithinCharacterLimit)
            {
                nameChecks.Add("isWithinCharacterLimit", isWithinCharacterLimit);
            }

            if (!isNotBlank)
            {
                nameChecks.Add("isNotBlank", isNotBlank);
            }

            if (!isNotReservedWord)
            {
                nameChecks.Add("isNotReservedWord", isNotReservedWord);
            }

            if (!isValidCharacters)
            {
                nameChecks.Add("isValidCharacters", isValidCharacters);
            }

            if (!isNotBeginWithInvalidCharacter)
            {
                nameChecks.Add("isNotBeginWithInvalidCharacter", isNotBeginWithInvalidCharacter);
            }

            return nameChecks.Count == 0;
        }

        private static bool IsNotBeginOrEndWithInvalidCharacter(string workSheetName)
        {
            return !workSheetName.StartsWith("\'") && !workSheetName.EndsWith("\'");
        }

        private static bool IsValidCharacters(string workSheetName)
        {
            Regex r = new Regex(@"[\[/\?\]\*\\\:]");

            return !r.IsMatch(workSheetName);
        }

        private static bool IsNotReservedWord(string workSheetName)
        {
            List<string> reservedWords = new List<string> { "history" };

            return !reservedWords.Contains(workSheetName.ToLower());
        }

        private static bool IsUnique(string workSheetName, IXLWorkbook workbook)
        {
            return !workbook.Worksheets.Contains(workSheetName);
        }

        private static bool IsWithinCharacterLimit(string workSheetName)
        {
            return workSheetName.Length <= 31;
        }

        private static string ModifyWorksheetName()
        {
            string workSheetName = "BHoM_Export_" + DateTime.Now.ToString("ddMMyy_HHmmss");

            return workSheetName;
        }
    }
}