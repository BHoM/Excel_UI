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
            string worksheetName = name;

            WorksheetValidation issue = IsValidName(worksheetName, workbook);
            while (issue != WorksheetValidation.Valid)
            {
                worksheetName = ModifyWorksheetName(worksheetName, issue);
                issue = IsValidName(worksheetName, workbook);
            }

            if (worksheetName != name)
            {
                BH.Engine.Base.Compute.RecordNote($"The selected name for worksheet {name} was not valid to Excel requirements. This has been renamed to {worksheetName}.");
            }

            return worksheetName;
        }

        /*************************************************/
        /****             Private Methods              ***/
        /*************************************************/

        private static WorksheetValidation IsValidName(string workSheetName, IXLWorkbook workbook)
        {
            if (!CheckIsUnique(workSheetName, workbook))
                return WorksheetValidation.NotUnique;

            if (string.IsNullOrWhiteSpace(workSheetName))
                return WorksheetValidation.Blank;

            if (!CheckIsNotReservedWord(workSheetName))
                return WorksheetValidation.ReservedWord;

            if (!CheckIsValidCharacters(workSheetName))
                return WorksheetValidation.NotValidCharacters;

            if (!CheckIsNotBeginOrEndWithInvalidCharacter(workSheetName))
                return WorksheetValidation.BeginOrEndWithInvalidCharacter;

            if (!CheckIsWithinCharacterLimit(workSheetName))
                return WorksheetValidation.NotWithinCharacterLimit;

            return WorksheetValidation.Valid;
        }

        /************ Checks *********************/

        private static bool CheckIsUnique(string workSheetName, IXLWorkbook workbook)
        {
            return !workbook.Worksheets.Contains(workSheetName);
        }

        private static bool CheckIsNotReservedWord(string workSheetName)
        {
            List<string> reservedWords = new List<string> { "history" };

            return !reservedWords.Contains(workSheetName.ToLower());
        }

        private static bool CheckIsValidCharacters(string workSheetName)
        {
            Regex r = new Regex(@"[\[/\?\]\*\\\:]");

            return !r.IsMatch(workSheetName);
        }

        private static bool CheckIsNotBeginOrEndWithInvalidCharacter(string workSheetName)
        {
            return !workSheetName.StartsWith("\'") && !workSheetName.EndsWith("\'");
        }

        private static bool CheckIsWithinCharacterLimit(string workSheetName)
        {
            return workSheetName.Length <= 31;
        }

        /************ Modifications *********************/
        private static string ModifyWorksheetName(string worksheetName, WorksheetValidation issue)
        {
            switch (issue)
            {
                case WorksheetValidation.ReservedWord:
                    return SetUniqueNameWithReservedName(worksheetName);

                case WorksheetValidation.Blank:
                    return SetGeneralName();

                case WorksheetValidation.NotUnique:
                    return SetUniqueName(worksheetName);

                case WorksheetValidation.BeginOrEndWithInvalidCharacter:
                    return RemoveInvalidBeginningOrEnding(worksheetName);

                case WorksheetValidation.NotValidCharacters:
                    return RemoveInvalidCharacters(worksheetName);

                case WorksheetValidation.NotWithinCharacterLimit:
                    return TrimName(worksheetName);

                default:
                case WorksheetValidation.Valid:
                    break;
            }

            return worksheetName;
        }

        private static string SetUniqueNameWithReservedName(string worksheetName)
        {
            return "BHoM_" + worksheetName;
        }

        private static string SetGeneralName()
        {
            return "BHoM_Export_" + DateTime.Now.ToString("ddMMyy_HHmmss");
        }

        private static string SetUniqueName(string worksheetName)
        {
            return DateTime.Now.ToString("ddMMyy HHmmssf") + "_" + worksheetName;
        }

        private static string RemoveInvalidBeginningOrEnding(string worksheetName)
        {
            if (worksheetName.StartsWith("\'"))
            {
                worksheetName = worksheetName.Substring(1);
            }

            if (worksheetName.EndsWith("\'"))
            {
                worksheetName = worksheetName.Substring(0, worksheetName.Length - 2);
            }

            return worksheetName;
        }

        private static string RemoveInvalidCharacters(string worksheetName)
        {
            List<char> invalidChars = new List<char>() { '[', '/', '?', ']', '*', '\\', ':' };

            invalidChars.ForEach(c => worksheetName = worksheetName.Replace(c.ToString(), String.Empty));
            return worksheetName;
        }

        private static string TrimName(string worksheetName)
        {
            if (worksheetName.Contains(" ") || worksheetName.Contains("-") || worksheetName.Contains("_") || worksheetName.Contains(","))
            {
                List<char> toBeRemovedChars = new List<char>() { ' ', '-', '_', ',' };
                toBeRemovedChars.ForEach(c => worksheetName = worksheetName.Replace(c.ToString(), String.Empty));
                return worksheetName;
            }

            return worksheetName.Substring(0, 31);
        }
    }
}
