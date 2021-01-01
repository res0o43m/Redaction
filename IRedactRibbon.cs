// Copyright (c) Microsoft Corporation.  All rights reserved.
using Microsoft.Office.Core;
using System.Drawing;

namespace Redaction
{
    public interface IRedactRibbon
    {
        void ButtonFindAndMarkClick(IRibbonControl control);
        void ButtonNextClick(IRibbonControl control);
        void ButtonPreviousClick(IRibbonControl control);
        void ButtonRedactClick(IRibbonControl control);
        void ButtonUnmarkAllClick(IRibbonControl control);
        void ButtonUnmarkClick(IRibbonControl control);
        string GetCustomUI(string RibbonID);
        string RibbonGetDescription(IRibbonControl control);
        bool RibbonGetEnabled(IRibbonControl control);
        string RibbonGetKeytip(IRibbonControl control);
        string RibbonGetLabel(IRibbonControl control);
        string RibbonGetScreentip(IRibbonControl control);
        string RibbonGetSupertip(IRibbonControl control);
        void RibbonLoad(IRibbonUI ribbonUI);
        Bitmap RibbonLoadImages(string image);
        void SplitButtonMarkClick(IRibbonControl control);
    }
}