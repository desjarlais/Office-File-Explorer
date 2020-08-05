/****************************** Module Header ******************************\
Module Name:  NoteSlideHelper.cs
Project:      Office File Explorer

Note Slide Helper for storing placeholder values from file

This source is subject to the following license.
See https://github.com/desjarlais/Office-File-Explorer/blob/master/LICENSE
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System;

namespace Office_File_Explorer.PowerPoint_Helpers
{
    public class NoteSlideHelper
    {
        public T2dHeader t2dHeader;
        public T2dDate t2dDate;
        public T2dSlideNumber t2dSlideNumber;
        public T2dSlideImage t2dSlideImage;
        public T2dPicture t2dPicture;
        public T2dFooter t2dFooter;
        public T2dNotes t2dNotes;
        public PresNotesSz pNotesSz;
    }

    public struct PresNotesSz
    {
        public Int64 Cx;
        public Int64 Cy;
    }

    public struct T2dHeader
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }

    public struct T2dDate
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }

    public struct T2dSlideNumber
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }

    public struct T2dPicture
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }

    public struct T2dFooter
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }

    public struct T2dNotes
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }

    public struct T2dSlideImage
    {
        public Int64 OffsetX;
        public Int64 OffsetY;
        public Int64 ExtentsCx;
        public Int64 ExtentsCy;
    }
}
