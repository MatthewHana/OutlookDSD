using System;
using System.ComponentModel;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookDSD
{
    [ComVisible(true)]
    public partial class OptionsPage : UserControl, PropertyPage
    {
        bool isDirty = false;
        bool SettingsLoaded = false;
        readonly BackgroundWorker UpdateCheckerBGWorker = new BackgroundWorker();
        string CurrentInstalledVersion = "";

        public OptionsPage()
        {
            // Create Form components
            InitializeComponent();
            // Suspend UI as we make changes
            SuspendLayout();
            // Load the version and check for updates
            Version_Load();
            // Load settings
            Settings_Load();
            //Resume UI
            ResumeLayout(false);
        }

        private void Settings_Load()
        {
            
            // Load each setting
            checkBox_ribbon_showOnEmail.Checked = Properties.Settings.Default.ribbon_showOnEmail;
            checkBox_ribbon_showOnExplorer.Checked = Properties.Settings.Default.ribbon_showOnExplorer;
            checkBox_bar_show.Checked = Properties.Settings.Default.bar_show;
            // Mark settings as loaded so that Change event will work
            SettingsLoaded = true;
        }

        private void Settings_Save()
        {
            Properties.Settings.Default.ribbon_showOnEmail = checkBox_ribbon_showOnEmail.Checked;
            Properties.Settings.Default.ribbon_showOnExplorer = checkBox_ribbon_showOnExplorer.Checked;
            Properties.Settings.Default.bar_show = checkBox_bar_show.Checked;

            Properties.Settings.Default.Save();
        }

        [DispId(-518)]
        public string Caption
        {
            get { return "OutlookDSD"; }
        }

        public void GetPageInfo(ref string HelpFile, ref int HelpContext)
        {
            return;
        }

        public void Apply()
        {
            // Cancel the UpdateChcker if we're closing the dialog
            if (UpdateCheckerBGWorker.IsBusy)
            {
                UpdateCheckerBGWorker.CancelAsync();
            }

            // If no changes have been made then just continue
            if (!isDirty)
            {
                return;
            }
            // Otherwise save the settings
            Settings_Save();
            MessageBox.Show("You may need to restart Oulook for your settings to take effect.");

            // Set isDirty back to false so that any other changes work accordingly
            isDirty = false;
        }

        public bool Dirty
        {
            get { return isDirty; }
            set { }
        }

        private void SettingChanged(object sender, EventArgs e)
        {
            // If we haven't finished loading the settings yet then just do nothing
            if (!SettingsLoaded)
            {
                return;
            }

            // If we've already made a change then don't continue any further.
            // This is to prevent GetPropertyPageSite() from being called over and over,
            // as it uses Reflection and can impact performance.
            if (isDirty)
            {
                return;
            }

            // User has toggled a setting change. Set the dirty flag
            isDirty = true;

            // Wrap it in a try because there is no gurantee that we are getting the PropertyPageSite
            try
            {
                // and let the parent page know so that it can activate the Apply button
                GetPropertyPageSite().OnStatusChange();
            }catch (System.Exception)
            {

            }
        }

        // Thanks to KyleMit at https://stackoverflow.com/a/21125886
        private PropertyPageSite GetPropertyPageSite()
        {
            Type objType = typeof(System.Object);
            string assemblyPath = objType.Assembly.CodeBase.Replace("mscorlib.dll", "System.Windows.Forms.dll").Replace("file:///", "");
            string assemblyName = System.Reflection.AssemblyName.GetAssemblyName(assemblyPath).FullName;

            Type unsafeNativeMethods = Type.GetType(System.Reflection.Assembly.CreateQualifiedName(assemblyName, "System.Windows.Forms.UnsafeNativeMethods"));
            Type oleObjectType = unsafeNativeMethods.GetNestedType("IOleObject");

            System.Reflection.MethodInfo methodInfo = oleObjectType.GetMethod("GetClientSite");
            Object propertyPageSite = methodInfo.Invoke(this, null);

            return (PropertyPageSite) propertyPageSite;
        }

#pragma warning disable IDE1006 // Naming Styles
        private void btn_Update_Click(object sender, EventArgs e)
#pragma warning restore IDE1006 // Naming Styles
        {
            // Open a URL to the GitHub page
            string url = "https://github.com/MatthewHana/OutlookDSD";
            Web_GoTo(url);
        }

#pragma warning disable IDE1006 // Naming Styles
        private void btn_Help_Click(object sender, EventArgs e)
#pragma warning restore IDE1006 // Naming Styles
        {
            // Open a URL to the GitHub page
            string url = "https://github.com/MatthewHana/OutlookDSD";
            Web_GoTo(url);
        }

        private void Web_GoTo(string url)
        {
            System.Diagnostics.Process.Start(url);
        }
        private void Version_Load()
        {
            // List the current version
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo versionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            CurrentInstalledVersion = versionInfo.ProductVersion.ToString();
            lbl_VersionInstalled.Text = CurrentInstalledVersion;

            // Start the UpdateChecker background 
            UpdateCheckerBGWorker.WorkerReportsProgress = true;
            UpdateCheckerBGWorker.DoWork += new DoWorkEventHandler(Version_CheckForUpdate);
            UpdateCheckerBGWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Version_UpdateAvaliable);
            UpdateCheckerBGWorker.RunWorkerAsync();
        }

        private void Version_UpdateAvaliable(object sender, RunWorkerCompletedEventArgs e)
        {
            // If there was an error, stop
            if(e.Error != null)
            {
                lbl_VersionAvaliable.Text = "Error";
                return;
            }
            string LatestVersion = (string) e.Result;

            // Make sure the versions string obtained is a valid one
            string validRegex = @"^(\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2})$";
            bool ValidVersion = Regex.IsMatch(LatestVersion, validRegex);
            if (!ValidVersion)
            {
                lbl_VersionAvaliable.Text = "Error";
                return;
            }

            lbl_VersionAvaliable.Text = LatestVersion;
            
            //Compare versions to see if there is an update
            var InstalledVerionObj = new Version(CurrentInstalledVersion);
            var LatestVerionObj = new Version(LatestVersion);
            bool UpdateAvaliable = InstalledVerionObj.CompareTo(LatestVerionObj) < 0;

            // If there's no update, stop here
            if (UpdateAvaliable == false)
            {
                return;
            }

            // Otherwise make UI changes
            btn_Update.Visible = true;
            lbl_VersionAvaliable.ForeColor = System.Drawing.Color.Red;
            lbl_VersionAvaliable.Font = new System.Drawing.Font(lbl_VersionAvaliable.Font, System.Drawing.FontStyle.Bold);
        }

        private void Version_CheckForUpdate(object sender, DoWorkEventArgs e)
        {
            //Set URL to check for updates
            string updateCheckURL = "https://raw.githubusercontent.com/MatthewHana/OutlookDSD/main/LATEST";

            //Enable HTTPS
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpWebRequest request = WebRequest.Create(updateCheckURL) as HttpWebRequest;
            request.Accept = "text/plain";
            request.UserAgent = Web_GetUA();

            // Get the response. We don't need to catch anything because we check for errors in our RunWorkerCompleted event callback
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            var encoding = ASCIIEncoding.ASCII;
            string responseText;
            using (var reader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
            {
                responseText = reader.ReadToEnd();
            }
            e.Result = responseText;
        }

        private string Web_GetUA()
        {
            string UA = "Mozilla/5.0 (Windows NT ";
            
            UA += System.Environment.OSVersion.Version.Major + "." + System.Environment.OSVersion.Version.Minor + ";";
            if (System.Environment.Is64BitOperatingSystem)
            {
                UA += " Win64; ";
            }
            else
            {
                UA += " Win32; ";
            }

            UA += System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString().ToLower() + ") ";
            UA += "OutlookDSD/" + CurrentInstalledVersion;

            return UA;
        }
    }
}
