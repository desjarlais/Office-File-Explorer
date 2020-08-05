/****************************** Module Header ******************************\
Module Name:  NumberingHelper.cs
Project:      Office File Explorer

ListTemplate Numbering Helper class

This source is subject to the following license.
See https://github.com/desjarlais/Office-File-Explorer/blob/master/LICENSE
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

namespace Office_File_Explorer.Word_Helpers
{
    public class NumberingHelper
    {
        public string NumFormat { get; set; }
        public int AbsNumId { get; set; }
        public int NumId { get; set; }
    }
}
