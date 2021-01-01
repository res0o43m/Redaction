using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Redaction.Properties;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon2();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Redaction
{
    [ComVisible(true)]
    public partial class Ribbon2 : Office.IRibbonExtensibility
    {
        private static Word.WdColor ShadingColor = (Word.WdColor)12697792; //marks are gray by default, but can be changed by setting this property.

        private object Missing = Type.Missing;
        private object CollapseStart = Word.WdCollapseDirection.wdCollapseStart;
        private object CollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;

        private Office.IRibbonUI ribbon;
        private Word.Application Application;
        private Word.ApplicationEvents4_WindowSelectionChangeEventHandler SelectionChangeEvent;
        private Word.ApplicationEvents4_DocumentChangeEventHandler DocumentChangeEvent;

        public Ribbon2()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Redaction.Ribbon2.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void RibbonLoad(Office.IRibbonUI ribbonUI)
        {

            this.ribbon = ribbonUI;
            //register events
            Application = Globals.ThisAddIn.Application;
            SelectionChangeEvent = new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Application.WindowSelectionChange += SelectionChangeEvent;
            DocumentChangeEvent = new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            Application.DocumentChange += DocumentChangeEvent;
        }

         public bool RibbonGetEnabled(Office.IRibbonControl control)
        {
            if (Application.Documents.Count == 0)
                return false;
            else if ((control.Id == "splitButtonMark" || control.Id == "splitButtonUnmark") && Application.Selection != null && Application.Selection.Type == Word.WdSelectionType.wdSelectionColumn)
                return false;
            else if (control.Id != "buttonMarkOfficeMenu" && Application.Selection != null && Application.Selection.StoryType == Word.WdStoryType.wdCommentsStory)
                return false;
            else if (Application.ActiveDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
                return false;
            else
                return true;
        }

        public string RibbonGetLabel(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "buttonMarkOfficeMenu":
                case "buttonRedact":
                    return Resources.RedactMenuItemLabel;
                case "groupRedact":
                    return Resources.RedactGroupLabel;
                case "buttonUnmark":
                case "splitButtonUnmark__btn":
                    return Resources.UnmarkLabel;
                case "buttonUnmarkAll":
                    return Resources.UnmarkAllLabel;
                case "buttonPrevious":
                    return Resources.PreviousLabel;
                case "buttonNext":
                    return Resources.NextLabel;
                case "buttonMark":
                case "splitButtonMark__btn":
                    return Resources.MarkLabel;
                case "buttonFindAndMark":
                    return Resources.FindAndMarkLabel;
                default:
                    Debug.Fail("unknown control requested a label: " + control.Id);
                    return null;
            }
        }

        public string RibbonGetDescription(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "buttonMarkOfficeMenu":
                    return Resources.RedactMenuItemDescription;
                default:
                    Debug.Fail("unknown control requested a description: " + control.Id);
                    return null;
            }
        }

        public string RibbonGetScreentip(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButtonMark__btn":
                case "buttonMark":
                    return Resources.MarkScreentip;
                case "buttonRedact":
                    return Resources.RedactScreentip;
                case "splitButtonUnmark__btn":
                case "buttonUnmark":
                    return Resources.UnmarkScreentip;
                case "buttonPrevious":
                    return Resources.PreviousScreentip;
                case "buttonNext":
                    return Resources.NextScreentip;
                default:
                    Debug.Fail("unknown control requested a screentip: " + control.Id);
                    return null;
            }
        }

        public string RibbonGetSupertip(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButtonMark__btn":
                case "buttonMark":
                    return Resources.MarkSupertip;
                case "splitButtonMark":
                    return Resources.MarkSplitMenuSupertip;
                case "buttonRedact":
                    return Resources.RedactSupertip;
                case "splitButtonUnmark__btn":
                case "buttonUnmark":
                    return Resources.UnmarkSupertip;
                case "splitButtonUnmark":
                    return Resources.UnmarkSplitMenuSupertip;
                case "buttonPrevious":
                    return Resources.PreviousSupertip;
                case "buttonNext":
                    return Resources.NextSupertip;
                case "buttonFindAndMark":
                    return Resources.FindAndMarkSupertip;
                default:
                    Debug.Fail("unknown control requested a supertip: " + control.Id);
                    return null;
            }
        }

        public string RibbonGetKeytip(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButtonMark":
                    return Resources.MarkKeytip;
                case "splitButtonUnmark":
                    return Resources.UnmarkKeytip;
                case "buttonPrevious":
                    return Resources.PreviousKeytip;
                case "buttonNext":
                    return Resources.NextKeytip;
                default:
                    Debug.Fail("unknown control requested a supertip: " + control.Id);
                    return null;
            }
        }

        public void SplitButtonMarkClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            MarkSelection();
            TurnOnEvents();
        }

        public void ButtonUnmarkClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            UnmarkSelection();
            TurnOnEvents();
        }

        public void ButtonUnmarkAllClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            UnmarkDocument();
            TurnOnEvents();
        }

        public void ButtonPreviousClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            SelectPreviousMark();
            TurnOnEvents();
        }

        public void ButtonNextClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            SelectNextMark();
            TurnOnEvents();
        }

        public void ButtonRedactClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            RedactDocument();
            TurnOnEvents();
        }

        public void ButtonFindAndMarkClick(Office.IRibbonControl control)
        {
            TurnOffEvents();
            FindAndMark();
            TurnOnEvents();
        }

        #endregion

        #region Word Events

        private void Application_WindowSelectionChange(Word.Selection selection)
        {
            InvalidateRedactionControls();
        }

        private void Application_DocumentChange()
        {
            InvalidateRedactionControls();
        }

        /// <summary>
        /// Invalidate the controls added by the redaction tool.
        /// </summary>
        private void InvalidateRedactionControls()
        {
            ribbon.InvalidateControl("splitButtonMark");
            ribbon.InvalidateControl("splitButtonUnmark");
            ribbon.InvalidateControl("buttonPrevious");
            ribbon.InvalidateControl("buttonNext");
        }

        private void TurnOffEvents()
        {
            try
            {
                Application.WindowSelectionChange -= SelectionChangeEvent;
                Application.DocumentChange -= DocumentChangeEvent;
            }
            catch (NullReferenceException) // if we fail to finish something, turn on never gets called, which would cause this to fail forever
            { }
        }

        private void TurnOnEvents()
        {
            Application.WindowSelectionChange += SelectionChangeEvent;
            Application.DocumentChange += DocumentChangeEvent;
            InvalidateRedactionControls();
        }

        #endregion
        #region Helpers
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
  
        #endregion
    }
}
