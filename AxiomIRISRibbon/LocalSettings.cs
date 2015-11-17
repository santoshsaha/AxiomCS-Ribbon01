using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace AxiomIRISRibbon
{
    public class LocalSettings
    {
        public enum Instances { Dev, IT, UAT, Prod }
        public enum Themes { Windows8, Dark, Office }

        private Instances ?inst;
        private bool debug;
        private bool ssologin;
        private bool showalllogins;

        private string soapversion;

        private string devurl;
        private string devorgid;
        private string iturl;
        private string itorgid;
        private string uaturl;
        private string uatorgid;
        private string produrl;
        private string prodorgid;

        private string reportsdevurl;
        private string reportsiturl;
        private string reportsuaturl;
        private string reportsprodurl;

        private Themes theme;
        private string themecolor;

        public Settings settingsDialog;

        public Instances ?Inst
        {
            get { return this.inst; }
            set { this.inst = value;

            if (this.settingsDialog != null)
            {
                this.settingsDialog.SetInstance(value);
            }
                
                SaveLocalSettings(); }            
        }

        public bool Debug
        {
            // Change Nov : Auto Login
          //   get { return true; }
             get { return this.debug; }
            set
            {
                this.debug = value;

                if (this.settingsDialog != null)
                {
                    this.settingsDialog.SetDebug(value);
                }
                SaveLocalSettings();
            }
        }

        public bool SSOLogin
        {
            get { return this.ssologin; }
            set
            {
                this.ssologin = value;

                if (this.settingsDialog != null)
                {
                    this.settingsDialog.SetSSOLogin(value);
                }
                SaveLocalSettings();
            }
        }

        public bool ShowAllLogins
        {
            get { return this.showalllogins; }
            set
            {
                this.showalllogins = value;

                if (this.settingsDialog != null)
                {
                    this.settingsDialog.SetShowAllLogins(value);
                }

                SaveLocalSettings();
            }
        }

        public string DevUrl
        {
            get { return this.devurl; }
            set { this.devurl = value; SaveLocalSettings(); }
        }

        public string DevOrgId
        {
            get { return this.devorgid; }
            set { this.devorgid = value; SaveLocalSettings(); }
        }

        public string ITUrl
        {
            get { return this.iturl; }
            set { this.iturl = value; SaveLocalSettings(); }
        }

        public string ITOrgId
        {
            get { return this.itorgid; }
            set { this.itorgid = value; SaveLocalSettings(); }
        }

        public string UATUrl
        {
            get { return this.uaturl; }
            set { this.uaturl = value; SaveLocalSettings(); }
        }

        public string UATOrgId
        {
            get { return this.uatorgid; }
            set { this.uatorgid = value; SaveLocalSettings(); }
        }

        public string ProdUrl
        {
            get { return this.produrl; }
            set { this.produrl = value; SaveLocalSettings(); }
        }

        public string ProdOrgId
        {
            get { return this.prodorgid; }
            set { this.prodorgid = value; SaveLocalSettings(); }
        }

        public string SoapVersion
        {
            get { return this.soapversion; }
            set { this.soapversion = value; SaveLocalSettings(); }
        }

        public string ReportsDevUrl
        {
            get { return this.reportsdevurl; }
            set { this.reportsdevurl = value; SaveLocalSettings(); }
        }

        public string ReportsITUrl
        {
            get { return this.reportsiturl; }
            set { this.reportsiturl = value; SaveLocalSettings(); }
        }

        public string ReportsUATUrl
        {
            get { return this.reportsuaturl; }
            set { this.reportsuaturl = value; SaveLocalSettings(); }
        }

        public string ReportsProdUrl
        {
            get { return this.reportsprodurl; }
            set { this.reportsprodurl = value; SaveLocalSettings(); }
        }

        public Themes Theme
        {
            get { return this.theme; }
            set
            {
                this.theme = value;
                if (this.settingsDialog != null)
                {
                    this.settingsDialog.SetTheme();
                }
                SaveLocalSettings();
            }
        }

        public string ThemeColor
        {
            get { return this.themecolor; }
            set
            {
                this.themecolor = value;
                if (this.settingsDialog != null)
                {
                    this.settingsDialog.SetTheme();
                }
                SaveLocalSettings();
            }
        }

        public LocalSettings()
        {
            //DevUrl = "https://dev-CS--IRIS----RKSB1-CS17-my-salesforce-com.iris-dev.rowini.net:29443",
            //    DevPartner = "/services/Soap/u/29.0/00Dg0000003Nc8C",
            //    DevMeta = "/services/Soap/m/29.0/00Dg0000003Nc8C"

            this.settingsDialog = null;

            this.inst = GetInstance("Inst", Instances.Prod);
            this.debug = GetBool("Debug",false);
            this.ssologin = GetBool("SSOLogin", false);
            this.showalllogins = GetBool("ShowAllLogins", false);
            this.soapversion = GetString("SoapVersion", "29.0");

            this.devurl = GetString("DevUrl", "https://dev-CS--IRIS----RKSB1-CS17-my-salesforce-com.iris-dev.rowini.net:29443");
            this.devorgid = GetString("DevOrgId", "00Dg0000003Nc8C");

            this.iturl = GetString("ITUrl", "");
            this.itorgid = GetString("ITOrgId", "");

            this.uaturl = GetString("UATUrl", "");
            this.uatorgid = GetString("UATOrgId", "");

            this.produrl = GetString("ProdUrl", "");
            this.prodorgid = GetString("ProdOrgId", "");

            this.reportsdevurl = GetString("ReportsDevUrl", "https://reprots.irisbyaxiom.com/SSOLogin");
            this.reportsiturl = GetString("ReportsITUrl", "https://reprots.irisbyaxiom.com/SSOLogin");
            this.reportsuaturl = GetString("ReportsUATUrl", "https://reprots.irisbyaxiom.com/SSOLogin");
            this.reportsprodurl = GetString("ReportsProdUrl", "https://reprots.irisbyaxiom.com/SSOLogin");


            this.theme = GetTheme("Theme", Themes.Windows8);
            this.themecolor = GetString("ThemeColor", "#DE5827");
        }

        public void Reset()
        {
            Properties.Settings.Default.Reset();
            this.inst = GetInstance("Inst", Instances.Prod);
            this.ssologin = GetBool("SSOLogin", false);
            this.debug = GetBool("Debug", false);
            this.showalllogins = GetBool("ShowAllLogins", false);
            this.soapversion = GetString("SoapVersion", "29.0");
            this.devurl = GetString("DevUrl", "https://dev-CS--IRIS----RKSB1-CS17-my-salesforce-com.iris-dev.rowini.net:29443");
            this.devorgid = GetString("DevOrgId", "00Dg0000003Nc8C");
            this.iturl = GetString("ITUrl", "");
            this.itorgid = GetString("ITOrgId", "");
            this.uaturl = GetString("UATUrl", "");
            this.uatorgid = GetString("UATOrgId", "");
            this.produrl = GetString("ProdUrl", "");
            this.prodorgid = GetString("ProdOrgId", "");

            this.reportsdevurl = GetString("ReportsDevUrl", "");
            this.reportsiturl = GetString("ReportsITUrl", "");
            this.reportsuaturl = GetString("ReportsUATUrl", "");
            this.reportsprodurl = GetString("ReportsProdUrl", "");

            this.theme = GetTheme("Theme", Themes.Windows8);
            this.themecolor = GetString("ThemeColor", "#DE5827");

        }

        public void SaveLocalSettings()
        {
            SetInstance("Inst", this.inst);
            SetBool("Debug", this.Debug);
            SetBool("SSOLogin", this.SSOLogin);
            SetBool("ShowAllLogins", this.ShowAllLogins);
            SetString("SoapVersion", "29.0");
            SetString("DevUrl", this.DevUrl);
            SetString("DevOrgId", this.DevOrgId);
            SetString("ITUrl", this.ITUrl);
            SetString("ITOrgId", this.ITOrgId);
            SetString("UATUrl", this.UATUrl);
            SetString("UATOrgId", this.UATOrgId);
            SetString("ProdUrl", this.ProdUrl);
            SetString("ProdOrgId", this.ProdOrgId);
            SetTheme("Theme", this.theme);
            SetString("ThemeColor", this.themecolor);

            SetString("ReportsDevUrl", this.ReportsDevUrl);
            SetString("ReportsITUrl", this.ReportsITUrl);
            SetString("ReportsUATUrl", this.ReportsUATUrl);
            SetString("ReportsProdUrl", this.ReportsProdUrl);

            Properties.Settings.Default.Save();
            
        }

        private bool GetBool(string name,bool deflt){
            if(PropertyExists(name))
            {
                return (bool)Properties.Settings.Default[name];
            }
            else
            {
                return deflt;
            }
        }

        private void SetBool(string name, bool val)
        {
            if (!PropertyExists(name))
            {
                
            }
            else
            {
                Properties.Settings.Default[name] = val;
            }
        }


        private string GetString(string name, string deflt)
        {
            if (PropertyExists(name))
            {
                return (string)Properties.Settings.Default[name];
            }
            else
            {
                return deflt;
            }
        }

        private void SetString(string name, string val)
        {
            if (!PropertyExists(name))
            {

            }
            else
            {
                Properties.Settings.Default[name] = val;
            }
        }

        private Instances GetInstance(string name, Instances deflt)
        {
            if (PropertyExists(name))
            {
                string inst = (string)Properties.Settings.Default[name];
                switch (inst)
                {
                    case "Dev":
                        return Instances.Dev;
                    case "IT":
                        return Instances.IT;
                    case "UAT":
                        return Instances.UAT;
                    default:
                        return Instances.Prod;
                }
            }
            else
            {
                return deflt;
            }
        }

        private void SetInstance(string name, Instances ?val)
        {
            if (!PropertyExists(name))
            {

            }
            else {
                string strval = "";
                switch (val)
                {
                    case Instances.Dev:
                        strval = "Dev";
                        break;
                    case Instances.IT:
                        strval = "IT";
                        break;
                    case Instances.UAT:
                        strval = "UAT";
                        break;
                    default:
                        strval = "Prod";
                        break;
                }
                Properties.Settings.Default[name] = strval;
            }           
        }

        private Themes GetTheme(string name, Themes deflt)
        {
            if (PropertyExists(name))
            {
                string inst = (string)Properties.Settings.Default[name];
                switch (inst)
                {
                    case "Windows8":
                        return Themes.Windows8;
                    case "Dark":
                        return Themes.Dark;
                    case "Office":
                        return Themes.Office;
                    default:
                        return Themes.Windows8;
                }
            }
            else
            {
                return deflt;
            }
        }

        private void SetTheme(string name, Themes? val)
        {
            if (!PropertyExists(name))
            {

            }
            else
            {
                string strval = "";
                switch (val)
                {
                    case Themes.Windows8:
                        strval = "Windows8";
                        break;
                    case Themes.Dark:
                        strval = "Dark";
                        break;
                    case Themes.Office:
                        strval = "Office";
                        break;
                    default:
                        strval = "Windows8";
                        break;
                }
                Properties.Settings.Default[name] = strval;
            }
        }

        private bool PropertyExists(string name)
        {            
            foreach (SettingsProperty p in Properties.Settings.Default.Properties)
            {
                if (p.Name == name)
                {
                    return true;
                }                
            }
            return false;
        }

    }
}
