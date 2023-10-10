using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OutlookDSD
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private readonly List<Office.IRibbonUI> ribbons = new List<Office.IRibbonUI> ();
        private readonly Dictionary<string, Validator> validatorCache = new Dictionary<string, Validator>();
        private readonly Dictionary<string, bool> enabledStatus = new Dictionary<string, bool>();

        public Ribbon()
        {
        }

        public void Invalidate()
        {
            foreach(Office.IRibbonUI ribbon in ribbons)
            {
                ribbon.Invalidate();
            }

        }
        public void CacheResults(Validator newValidatorObj)
        {
            MailItem mailItem = newValidatorObj.GetMailItem();
            string entryID = mailItem.EntryID;
            if (!validatorCache.ContainsKey(entryID))
            {
                validatorCache.Add(entryID, newValidatorObj);
            }
            Invalidate();
        }

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            string xmlFile;
            string settingCheck;
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    settingCheck = "ribbon_showOnExplorer";
                    xmlFile = "OutlookDSD.RibbonMain.xml";
                    break;
                case "Microsoft.Outlook.Mail.Read":
                    settingCheck = "ribbon_showOnEmail";
                    xmlFile = "OutlookDSD.RibbonMessage.xml";
                    break;
                default:
                    return null;
            }
            bool showRibbon = (bool) Properties.Settings.Default[settingCheck];
            if(showRibbon == false) {
                return null;
            }
            string xml = GetResourceText(xmlFile);
            return xml;
        }

            #endregion

            #region Ribbon Callbacks
            // Function signature for all callbacks is at https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)

            // Hanldle clicking on a mechnaism button
            public void Btn_Click(Office.IRibbonControl control){

            // Get the current email from the context that the control belongs to
            MailItem mailItem = GetMailItemFromCurrntControl(control);
            if (mailItem == null)
            {
                return;
            }

            // Get the result of the mechnaism clicked on
            string mechanism = ControlIDtoMechanism(control)[0].ToLower();
            string mechanismUpper = mechanism.ToUpper();
            Validator validator = ValidatorGet(mailItem);
            string[] result = validator.Results()[mechanism];
            
            // Set the details string to the full details of the result summary
            string details = result[1];

            //If the mechanism result was none, then set an error message as the details string
            if (result[0] == Validator.RESULT_NONE)
            {
                details = mechanismUpper + " was not used.";
            }

            // Display the messagebox to the user.
            MessageBox.Show(details, "Details for " + mechanismUpper + " - OutlookDSD");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbons.Add(ribbonUI);
        }
        #endregion

        public bool Btn_getEnabled(Office.IRibbonControl ctrl)
        {
            MailItem mailitem = GetMailItemFromCurrntControl(ctrl);
            if(mailitem != null)
            {
                if (enabledStatus.ContainsKey(mailitem.EntryID))
                {
                    return enabledStatus[mailitem.EntryID];
                }
            }

            return false;
        }
        #region Helpe
        
        public System.Drawing.Image Btn_GetImage(Office.IRibbonControl ctrl)
        {
            // Get the current email from the context that the control belongs to
            MailItem mailItem = GetMailItemFromCurrntControl(ctrl);

            // If the controls are disabled then return the blank image called disabled
            if (null == mailItem || enabledStatus.ContainsKey(mailItem.EntryID) == false || enabledStatus[mailItem.EntryID] == false)
            {
                return (System.Drawing.Image)Properties.Resources.ResourceManager.GetObject("disabled");
            }

            string[] labelInfo = ControlIDtoMechanism(ctrl);
            string mechnaism = labelInfo[0].ToLower();
            string IconName;

            Validator validator = ValidatorGet(mailItem);

            if (!validator.Results().ContainsKey(mechnaism))
            {
                IconName = "error";
            }
            else
            {
                IconName = validator.Results()[mechnaism][0];
            }
            return (System.Drawing.Image) Properties.Resources.ResourceManager.GetObject(IconName);
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

        private Validator ValidatorGet(MailItem mailItem)
        {
            string entryID = mailItem.EntryID;
            if (validatorCache.ContainsKey(entryID))
            {
                validatorCache.TryGetValue(entryID, out Validator validator);
                return validator;
            }
            return new Validator(null);
        }

        private string[] ControlIDtoMechanism(Office.IRibbonControl control)
        {
            string controlId = control.Id;
            if (!controlId.StartsWith("emailValidation")){
                return new string[] { String.Empty, String.Empty };
            }
            string[] nameParts = controlId.Split("_"[0]);
            return new string[]{
                nameParts[2],
                nameParts[1],
            };
        }

        private MailItem GetMailItemFromCurrntControl(Office.IRibbonControl control)
        {
            int contextClass = control.Context.Class;

            if (contextClass == 35) // Email - IPM.Note
            {
                return control.Context.CurrentItem;
            }
            else if (contextClass == 34) // Outlook Explorer
            {
                // Wrap the whole thing in a try/catch statement in case ActiveExplorer throws an exception
                try
                {
                    var selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
                    if (selection.Count == 1) // One email is selected, let's parse it
                    {
                        MailItem emailItem = selection[1]; // Not a bug - Selection index starts at 1 and not 0!
                        return emailItem;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch(System.Exception)
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public void Enable(MailItem emailItem)
        {
            SetEnabled(true, emailItem);
        }

        public void Disable(MailItem emailItem)
        {
            SetEnabled(false, emailItem);
        }

        private void SetEnabled(bool isEnabled, MailItem emailItem)
        {
            if(emailItem != null)
            {
                enabledStatus[emailItem.EntryID] = isEnabled;
            }
            Invalidate();
        }

        #endregion
    }
}
