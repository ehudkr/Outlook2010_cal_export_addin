using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Outlook2010_cal_export_adding
{
    [ComVisible(true)]
    public class context_menu_export : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new context_menu_export();
        }

        public context_menu_export()
        {
            //control.Context.
            //Outlook2010_cal_export_adding.Globals.ThisAddIn.Application.
            
            //Outlook.CalendarSharing cls = (Outlook.CalendarSharing) 
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Outlook2010_cal_export_adding.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void ExportButtonClick(Office.IRibbonControl control)
        {
            
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

        //public Image GetIcon(Office.IRibbonControl control)
        //{
        //    return Outlook2010_cal_export_adding.Properties.Resources.icon;
        //}

        //public string GetSynchronisationLabel(Office.IRibbonControl control)
        //{
        //    return "Synchronize";
        //}

        //public void ShowMessageClick(Office.IRibbonControl control)
        //{
        //    System.Windows.Forms.MessageBox.Show("You've clicked the synchronize context menu item", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //}
    }
}
