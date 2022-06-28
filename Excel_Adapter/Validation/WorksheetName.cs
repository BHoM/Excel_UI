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
using DocumentFormat.OpenXml.Spreadsheet;
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
            string worksheetName = name;

            List<WorksheetValidation> validationIssues = new List<WorksheetValidation>();

            bool isValidName = IsValidName(worksheetName, workbook, out validationIssues);
            while (!isValidName)
            {
                worksheetName = ModifyWorksheetName(worksheetName, validationIssues, workbook);
                isValidName = IsValidName(worksheetName, workbook, out validationIssues);
            }


            if (worksheetName != name)
            {
                BH.Engine.Base.Compute.RecordError("Name of worksheet has been adjusted to a name that which is compatible with Excel naming limitations.");
            }

            return worksheetName;
        }


        /***************************************************/

        private static string ModifyWorksheetName(string worksheetName, List<WorksheetValidation> validationIssues, IXLWorkbook workbook)
        {
            switch (validationIssues[0])
            {
                case WorksheetValidation.isReservedWord:
                case WorksheetValidation.isBlank:
                    return SetGeneralName();

                case WorksheetValidation.isNotUnique:
                    return SetUniqueName(workbook, worksheetName);

                case WorksheetValidation.isBeginOrEndWithInvalidCharacter:
                    return RemoveInvalidBeginningOrEnding(worksheetName);

                case WorksheetValidation.isNotValidCharacters:
                    return RemoveInvalidCharacters(worksheetName);

                case WorksheetValidation.isNotWithinCharacterLimit:
                    return TrimName(worksheetName);

                default:
                    break;
            }

            return worksheetName;
        }

        private static string TrimName(string worksheetName)
        {
            if (worksheetName.ToString().Contains(" "))
            {
                worksheetName = worksheetName.Replace(" ", "");
                return worksheetName;
            }

            return worksheetName.Substring(0, 31);

        }

        private static string RemoveInvalidCharacters(string worksheetName)
        {
            List<char> InvalidChars = new List<char>() { '[', '/', '?', ']', '*', '\\', ':' };

            InvalidChars.ForEach(c => worksheetName = worksheetName.Replace(c.ToString(), String.Empty));
            return worksheetName;
        }

        private static string RemoveInvalidBeginningOrEnding(string worksheetName)
        {
            if (worksheetName.Substring(0, 1) == "\'")
            {
                worksheetName = worksheetName.Replace("\'", "");
            }

            if (worksheetName.Substring(worksheetName.Length - 1) == "\'")
            {
                worksheetName = worksheetName.Replace("\'", "");
            }

            return worksheetName;
        }

        private static string SetUniqueName(IXLWorkbook workbook, string worksheetName)
        {
            return DateTime.Now.ToString("ddMMyy HHmmssf") + "_" + worksheetName;
        }

        private static bool IsValidName(string workSheetName, IXLWorkbook workbook, out List<WorksheetValidation> validationIssues)
        {
            validationIssues = new List<WorksheetValidation>();

            if (!CheckIsUnique(workSheetName, workbook))
                validationIssues.Add(WorksheetValidation.isNotUnique);

            if (!CheckIsNullOrWhitespace(workSheetName))
                validationIssues.Add(WorksheetValidation.isBlank);

            if (!CheckIsNotReservedWord(workSheetName))
                validationIssues.Add(WorksheetValidation.isReservedWord);

            if (!CheckIsValidCharacters(workSheetName))
                validationIssues.Add(WorksheetValidation.isNotValidCharacters);

            if (!CheckIsNotBeginOrEndWithInvalidCharacter(workSheetName))
                validationIssues.Add(WorksheetValidation.isBeginOrEndWithInvalidCharacter);

            if (!CheckIsWithinCharacterLimit(workSheetName))
                validationIssues.Add(WorksheetValidation.isNotWithinCharacterLimit);

            if (validationIssues.Count > 0)
            {
                return false;
            }

            return true;
        }

        private static bool CheckIsNullOrWhitespace(string workSheetName)
        {
            return !string.IsNullOrWhiteSpace(workSheetName);
        }

        private static bool CheckIsNotBeginOrEndWithInvalidCharacter(string workSheetName)
        {
            return !workSheetName.StartsWith("\'") && !workSheetName.EndsWith("\'");
        }

        private static bool CheckIsValidCharacters(string workSheetName)
        {
            Regex r = new Regex(@"[\[/\?\]\*\\\:]");

            return !r.IsMatch(workSheetName);
        }

        private static bool CheckIsNotReservedWord(string workSheetName)
        {
            List<string> reservedWords = new List<string> { "history" };

            return !reservedWords.Contains(workSheetName.ToLower());
        }

        private static bool CheckIsUnique(string workSheetName, IXLWorkbook workbook)
        {
            return !workbook.Worksheets.Contains(workSheetName);
        }

        private static bool CheckIsWithinCharacterLimit(string workSheetName)
        {
            return workSheetName.Length <= 31;
        }

        private static string SetGeneralName()
        {
            return "BHoM_Export_" + DateTime.Now.ToString("ddMMyy_HHmmss");
        }

        private enum WorksheetValidation
        {
            isNotUnique,
            isNotWithinCharacterLimit,
            isBlank,
            isReservedWord,
            isNotValidCharacters,
            isBeginOrEndWithInvalidCharacter
        }
    }

}
