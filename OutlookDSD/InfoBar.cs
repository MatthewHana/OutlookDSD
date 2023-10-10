using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookDSD
{
    partial class InfoBar
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("OutlookDSD.InfoBar")]
        public partial class InfoBarFactory
        {

            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void InfoBarFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                Outlook.MailItem emailItem = (Outlook.MailItem)e.OutlookItem;

                if (emailItem == null)
                {
                    e.Cancel = true;
                    return;
                }
            }
        }

        #endregion

        private void InfoBar_FormRegionShowing(object sender, System.EventArgs e)
        {
            Outlook.MailItem emailItem = (Outlook.MailItem) OutlookItem;

            if (emailItem == null)
            {
                OutlookFormRegion.Visible = false;
                return;
            }

            // Only show on emails.
            if (emailItem.MessageClass != "IPM.Note")
            {
                OutlookFormRegion.Visible = false;
                return;
            }

            // Don't show for any draft emails.
            if(emailItem.Sent == false)
            {
                OutlookFormRegion.Visible = false;
                return;
            }

            // Now we check our settings to see if the user wants the bar to show
            string settingCheck = "bar_show";
            bool showBar = (bool) Properties.Settings.Default[settingCheck];
            if (showBar == false)
            {
                OutlookFormRegion.Visible = false;
                return;
            }

            // We don't need to worry about multiple messages being selected in explorer mode
            // because Outlook only displays one anyway.

            Validator validator = new Validator(emailItem);
            Dictionary<string, string[]> results = validator.Results();

            List<string> failed = new List<string>();
            List<string> missing = new List<string>();
            List<string> error = new List<string>();
            List<string> pass = new List<string>();

            // Iterate through each mechanism in result and add them to the appropriate list based on their result
            foreach (string mechanism in results.Keys)
            {
                string mechanismResult = results[mechanism][0];
                string mechanismUpper = mechanism.ToUpper();
                switch (mechanismResult)
                {
                    case Validator.RESULT_PASS:
                        pass.Add(mechanismUpper);
                        break;
                    case Validator.RESULT_FAIL:
                        failed.Add(mechanismUpper);
                        break;
                    case Validator.RESULT_NONE:
                        missing.Add(mechanismUpper);
                        break;
                    default:
                    case Validator.RESULT_ERROR:
                        error.Add(mechanismUpper);
                        break;

                }
            }

            int badMechanismsCount = error.Count + failed.Count + missing.Count;

            // If all mechnaisms are good then don't display anything.
            if (badMechanismsCount == 0)
            {
                OutlookFormRegion.Visible = false;
                return;
            }

            // Otherwise we're going to show the warning bar
            OutlookFormRegion.Visible = true;


            List<string> messages = new List<string>();
            if(failed.Count > 0)
            {
                messages.Add("This email has failed " + Helper.ListToSentence(failed) + " validation.");
            }
            if (missing.Count > 0)
            {
                messages.Add("This email is missing " + Helper.ListToSentence(missing) + " validation.");
            }
            if(error.Count > 0)
            {
                messages.Add("There was an error analysing " + Helper.ListToSentence(error) + " validation.");
            }
            messageLabel.Text = String.Join(" ", messages.ToArray());
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void InfoBar_FormRegionClosed(object sender, System.EventArgs e)
        {
        }

    }
}
