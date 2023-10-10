using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookDSD
{
    public partial class ThisAddIn
    {

        private Inspectors inspectors;
        private Explorers explorers;
        private Explorer activeExplorer;

        private Ribbon ribbon;

        private void AddIn_Startup(object sender, System.EventArgs e)
        {
            Register_Events();
        }

        private void AddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Register_Events()
        {
            // These Event Handlers need to be attached to objects that do not get garbage collected
            // therefore we need to set them as class level variables.
            inspectors = Application.Inspectors;
            explorers = Application.Explorers;
            activeExplorer = Application.ActiveExplorer();

            // This event handles opening email viewing window
            inspectors.NewInspector += new InspectorsEvents_NewInspectorEventHandler(Event_EmailOpened);

            // This event handles clicking on an email in the Otlook folder/explorer view
            activeExplorer.SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(Event_ExplorerSelected);

            // This event ensures that new explorer windows are also captured by this add-in
            explorers.NewExplorer += new ExplorersEvents_NewExplorerEventHandler(Event_NewExplorer);

            // Add the Add-in options dialog
            Application.OptionsPagesAdd += new ApplicationEvents_11_OptionsPagesAddEventHandler(Event_AddOptionsPage);
        }

        private void Event_NewExplorer(Explorer Explorer)
        {
            // This event handles clicking on an email in the Otlook folder/explorer view
            Explorer.SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(Event_ExplorerSelected);
        }

        private void Event_AddOptionsPage(PropertyPages Pages)
        {
            Pages.Add(new OptionsPage(), "OutlookDSD");    
        }

        private void Event_ExplorerSelected()
        {
            var selection = Application.ActiveExplorer().Selection;
            if (selection.Count == 1) // One email is selected, let's parse it
            {
                MailItem emailItem = selection[1]; // Not a bug. Selection index starts at 1 and not 0!
                Process_Email(emailItem);
            }
            else // Less or More than 1 email is selected. Simply toggle a redraw to have the controls redrawn as being disabled
            {
                ribbon.Invalidate();
                return;
            }
        }

        private void Event_EmailOpened(Inspector Inspector)
        {
            MailItem emailItem = Inspector.CurrentItem;
            Process_Email(emailItem);
        }

        private void Process_Email(MailItem emailItem)
        {
            // Stop if the mailItem is null
            if (emailItem == null)
            {
                return;
            }

            // Stop if there is no entry ID
            if (emailItem.EntryID == null)
            {
                return;
            }
            //Stop if the item is NOT an email
            if (emailItem.MessageClass != "IPM.Note")
            {
                return;
            }
            // Don't process any draft emails.
            if (emailItem.Sent == false)
            {
                ribbon.Disable(emailItem);
                return;
            }

            // Parse the email
            Validator validatorObj = new Validator(emailItem);

            ribbon.CacheResults(validatorObj);

            if (validatorObj.isSent)
            {
                ribbon.Disable(emailItem);
            }
            else
            {
                ribbon.Enable(emailItem);
            }

        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon();
            return ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(AddIn_Startup);
            Shutdown += new System.EventHandler(AddIn_Shutdown);
        }
        
        #endregion
    }
}
