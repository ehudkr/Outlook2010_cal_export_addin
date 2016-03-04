using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

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
            // Get the current view of the calendar, in order to bound the exported calendar items to that view.
            DateTime dateStart;
            DateTime dateEnd;
            Outlook.Explorer expl = Outlook2010_cal_export_adding.Globals.ThisAddIn.Application.ActiveExplorer();
            Outlook.View view = expl.CurrentView as Outlook.View;
            if (view.ViewType == Outlook.OlViewType.olCalendarView)
            {
                Outlook.CalendarView calView = view as Outlook.CalendarView;
                dateStart = calView.SelectedStartTime;
                dateEnd = calView.SelectedEndTime;
                //DateTime end = start.AddDays(7); - use only a click on Sunday and export one week automatically ?
                //Automate date time from sunday midnight to saturday 23:59 ?
                //Set the hours of properly - full from sunday midnight to saturday 23:59? or leave it to the user?
            }
            else
            {
                    // No dates were correctly marked. set date for a week, beginning sunday midight to saturday 23:59:

                // Set start date:
                if ((int)DateTime.Today.DayOfWeek == 1)
                {
                    // today is sunday, export from now and on:
                    dateStart = DateTime.Today;
                }
                else
                {
                    // today is not sunday, export from the nearest sunday and onwards:
                    // calculate the days untill sunday:
                    int daysUntilSunday = ((int)DayOfWeek.Sunday - (int)DateTime.Today.DayOfWeek + 7) % 7;
                    dateStart = DateTime.Today.AddDays(daysUntilSunday);
                }

                // Set end date:
                dateEnd = dateStart.AddDays(7);

                // since we choose the date, we set the hours correctlly: 
                // from sunday midnight to saturday 23:59
                dateStart.AddHours(00).AddMinutes(00).AddSeconds(00);
                dateEnd.AddHours(23).AddMinutes(59).AddSeconds(59);
            }

                    // Access the calendar and export events:

            Outlook.Folder calFolder = 
                Outlook2010_cal_export_adding.Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) 
                as Outlook.Folder;
            // Get the items in the wanted range of dates:
            Outlook.Items rangeAppts = GetAppointmentsInRange(calFolder, dateStart, dateEnd);
            // Iterate over wanted (filtered) events (appointment-items) and write them to file:
            if (rangeAppts != null)
            {
                var csvWriter = new StringBuilder();
                // Header format by Google Calendar: https://support.google.com/calendar/answer/37118?hl=en
                var headerLine = string.Format("{0},{1},{2}.{3},{4},{5},{6},{7},{8},{9}",
                                               "Subject", "Start date", "Start time", "End date", "End time",
                                               "All day event", "Description", "Location", "Private");
                csvWriter.AppendLine(headerLine);
                foreach (Outlook.AppointmentItem item in rangeAppts)
                {
                    // Outputs event's wanted properties:
                    var newLine = string.Format("{0},{1},{2}.{3},{4},{5},{6},{7},{8},{9}",
                                                item.Subject.ToString(), 
                                                item.Start.Date.ToString(), item.Start.TimeOfDay.ToString(),
                                                item.End.Date.ToString(), item.End.TimeOfDay.ToString(), 
                                                item.AllDayEvent.ToString(), item.FormDescription.ToString(), 
                                                item.Location, "");
                    csvWriter.AppendLine(newLine);
                    //Debug.WriteLine("Subject: " + appt.Subject + " Start: " + appt.Start.ToString("g"));
                }

                // Write to file:
                string desktopath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopath = desktopath + @"\Agenda " + dateStart.Date.ToString() + ".csv";
                File.WriteAllText(desktopath, csvWriter.ToString());
                //File.AppendAllText(somepath, csvWriter.ToString());
            }



            ////control.Context.
            ////Outlook2010_cal_export_adding.Globals.ThisAddIn.Application.

            ////Outlook.CalendarSharing cls = (Outlook.CalendarSharing) 

            //// Iterate through all the events (appointment-items) in the calendar
            //Outlook.Application oApp = new Outlook.Application();
            //Outlook.NameSpace mapiNamespace = oApp.GetNamespace("MAPI"); ;
            //Outlook.MAPIFolder CalendarFolder = mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            //Outlook.Items outlookCalendarItems = CalendarFolder.Items;

            //outlookCalendarItems.Sort("Start");
            ////include recurring events (just a "normal" events i.e once per time it occurs)
            ////Have to Sort by Start before setting IncludeRecurrences.
            //outlookCalendarItems.IncludeRecurrences = true;
            //foreach (Outlook.AppointmentItem item in outlookCalendarItems)
            //{
            //    if (item.Start >= dateStart && item.End <= dateEnd)
            //    {

            //    }
            //}
        }

        #endregion

        #region Helpers

        private Outlook.Items GetAppointmentsInRange(Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {
            // Create a filter based on start-and-end-dates.
            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";
            //Debug.WriteLine(filter);
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }

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
