/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2018, the respective contributors. All rights reserved.
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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.UI.Excel
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
