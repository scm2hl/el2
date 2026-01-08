using Prism.Events;
using System;
using System.Configuration;
using System.IO;
using System.Reflection.Metadata;
using System.Threading;
using System.Windows;


namespace El2Core.Services
{
    public interface IUserSettingsService
    {
        string ExplorerPath { get; set; }
        string PersonalFolder { get; set; }
        bool IsAutoSave { get; set; }
        bool IsSaveMessage { get; set; }
        bool IsRowDetails { get; set; }
        string Theme { get; set; }
        double FontSize { get; set; }
        string PlanedSetup { get; set; }
        double SizePercent { get; set; }
        int KWReview {  get; set; }
        int EmployTimeFormat { get; set; }
        bool IsDefaults();
        bool IsChanged { get; }
        void Save();
        void Reset();
        void Reload();
        void Upgrade();
    }
    public class UserSettingsService : IUserSettingsService
    {
        public string ExplorerPath
        {
            get { return Properties.Settings.Default.ExplorerPath; }
            set { Properties.Settings.Default[nameof(ExplorerPath)] = value; _isChanged = true; }
        }
        public string PersonalFolder
        {
            get { return Properties.Settings.Default.PersonalFolder; }
            set { Properties.Settings.Default[nameof(PersonalFolder)] = value; _isChanged = true; }
        }
        public bool IsAutoSave
        {
            get { return Properties.Settings.Default.IsAutoSave; }
            set { Properties.Settings.Default[nameof(IsAutoSave)] = value; _isChanged = true; }
        }
        public bool IsSaveMessage
        {
            get { return Properties.Settings.Default.IsSaveMessage; }
            set { Properties.Settings.Default[nameof(IsSaveMessage)] = value; _isChanged = true; }
        }
        public bool IsRowDetails
        {
            get { return Properties.Settings.Default.IsRowDetails;  }
            set { Properties.Settings.Default[nameof(IsRowDetails)] = value; _isChanged = true; }
        }
        public double FontSize
        {
            get { return Properties.Settings.Default.FontSize; }
            set { Properties.Settings.Default[nameof(FontSize)] = value; _isChanged = true; }
        }
        public double SizePercent
        {
            get { return Properties.Settings.Default.SizePercent; }
            set { Properties.Settings.Default[nameof(SizePercent)] = value; _isChanged = true; }
        }
        public string Theme
        {
            get { return Properties.Settings.Default.Theme; }
            set
            {
                if (Theme != value)
                {
                    Properties.Settings.Default[nameof(Theme)] = value; _isChanged = true;
                }
            }
        }

        public bool UpgradeFlag
        {
            get => Properties.Settings.Default.UpgradeFlag;
            set { Properties.Settings.Default[nameof(UpgradeFlag)] = value; }
        }
        private bool _isChanged;
        public bool IsChanged { get { return _isChanged; } }
        public string PlanedSetup
        {
            get => Properties.Settings.Default.PlanedSetup;
            set { Properties.Settings.Default[nameof(PlanedSetup)] = value; _isChanged = true; }
        }

        public int KWReview
        {
            get => Properties.Settings.Default.KWReview;
            set { Properties.Settings.Default[nameof(KWReview)] = value; _isChanged = true; }
        }

        public int EmployTimeFormat
        {
            get => Properties.Settings.Default.EmployTimeFormat;
            set { Properties.Settings.Default[nameof(EmployTimeFormat)] = value; _isChanged = true; }
        }

        public void Save()
        {
             Properties.Settings.Default.Save();
            _isChanged = false;
        }
        public void Reset()
        {           
            Properties.Settings.Default.Reset();
            _isChanged = false;
        }

        public void Reload()
        {
            Properties.Settings.Default.Reload();
        }

        public void Upgrade()
        {
            var fp = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;
            bool.TryParse(Environment.GetEnvironmentVariable("ClickOnce_IsNetworkDeployed"), out bool isNetworkDeployed);
            if (isNetworkDeployed)
            {
                var previous = Environment.GetEnvironmentVariable("EL2_PREVIOUS_VERSION_CONFIG", EnvironmentVariableTarget.User);

                if (previous != null)
                {
                    var curFileInfo = new FileInfo(fp);
                    var preFileInfo = new FileInfo(previous);
                    if (preFileInfo.Exists)
                    {
                        if (curFileInfo.DirectoryName != null)
                        {
                            var UrlhashNew = curFileInfo.Directory.Parent.Name;
                            var UrlhashOld = preFileInfo.Directory.Parent.Name;
                            var newFile = previous.Replace(UrlhashOld, UrlhashNew);
                            var newDir = new FileInfo(newFile).Directory;
                            if (newDir?.Exists == false) { Directory.CreateDirectory(newDir.FullName); }
                            if (previous.Equals(newFile) == false)
                                File.Copy(previous, newFile, true);
                        }
                    }
                }
                Environment.SetEnvironmentVariable("EL2_PREVIOUS_VERSION_CONFIG", fp, EnvironmentVariableTarget.User);
            }
            if (Properties.Settings.Default.UpgradeFlag == true)
            {
                try
                {
                    
                    Properties.Settings.Default.Upgrade();
                    Properties.Settings.Default.Reload();
   
                    UpgradeFlag = false;
                    Save();
                }
                catch (System.Exception e)
                {

                    MessageBox.Show(string.Format("{0}\n{1}", e.Message, e.InnerException), "Upgrade", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        bool IUserSettingsService.IsDefaults()
        {
            return false;
        }
    }
}
