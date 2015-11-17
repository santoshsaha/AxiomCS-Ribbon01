using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Telerik.Windows.Controls;
using Telerik.Windows.Controls.Navigation;
using System.ComponentModel;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for LoginSSO.xaml
    /// </summary>
    public partial class LoginSSO : RadWindow
    {
        private Data _d;
        private SForceEdit.Settings _settings;
        private LocalSettings _local;

        private LocalSettings.Instances ?_instance;

        private string _instancename;
        private string _url;
        private string _orgid;
        private string _soapversion;

        System.Windows.Forms.WebBrowser _webBrowser1;

        public LoginSSO()
        {
            // Login to Support CS SSO - just use a browser control to hit the SSO page and then
            // sniff for the Session Id and use that to login to the API

            // started with the WPF browser control but was very buggy! searching the internet said
            // the windows forms one was more stable, switched and seems to be fine

            // TODO - timeout detection - when salesforce timesout, there is a javascript error from the control
            // I think when SForce is trying to display the "You are about to be logged out due to inactivity"
            // this should be trapped and a message shown to the user with a better logout message/process

            InitializeComponent();            
            Utility.setTheme(this);

            RadWindowInteropHelper.SetAllowTransparency(this, false);

            _d = Globals.ThisAddIn.getData();
            _settings = Globals.ThisAddIn.getSettings();
            _local = Globals.ThisAddIn.GetLocalSettings();

            // try the windows form browser
            _webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.wfh1.Child = _webBrowser1;            
            _webBrowser1.Navigated += _webBrowser1_Navigated;
            _webBrowser1.DocumentCompleted += _webBrowser1_DocumentCompleted;

            // stop the reminders - salesforce puts up a reminder window - it doesn't work though
            // we get an IE window with an error so just cancel it 
            _webBrowser1.NewWindow += _webBrowser1_NewWindow;
            
        }

        void _webBrowser1_NewWindow(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
        }

        void _webBrowser1_DocumentCompleted(object sender, System.Windows.Forms.WebBrowserDocumentCompletedEventArgs e)
        {
            System.Windows.Forms.HtmlDocument doc = (System.Windows.Forms.HtmlDocument)_webBrowser1.Document;
            if (doc.Title.ToLower().Contains("error") || doc.Title.ToLower().Contains("if you can see this page"))
            {
                bsyInd.IsBusy = false;
                lblLoginMessage.Text = "There has been a problem with the SSO login - you need to be setup as a user in IRIS Core to access the ribbon.";
            }
        }


        void _webBrowser1_Navigated(object sender, System.Windows.Forms.WebBrowserNavigatedEventArgs e)
        {
            string url = e.Url.ToString();
            Globals.Ribbons.Ribbon1.SFDebug("Navigated", url);

            // lblLoginMessage.Text = "";

            System.Windows.Forms.HtmlDocument doc = (System.Windows.Forms.HtmlDocument)_webBrowser1.Document;

            if (doc.Url.ToString().StartsWith("res://ieframe.dll"))
            {
                bsyInd.IsBusy = false;
                lblLoginMessage.Text = "There has been a problem with the SSO login";
            }

            if (doc.Title.ToLower().Contains("error") || doc.Title.ToLower().Contains("if you can see this page"))
            {
                bsyInd.IsBusy = false;
                lblLoginMessage.Text = "There has been a problem with the SSO login - you need to be setup as a user in IRIS Core to access the ribbon.";
            }

            if (e.Url.AbsolutePath.EndsWith("home.jsp"))
            {
                // get the session id from the cookie
                string sid = "";
                string cookies = "";
                try
                {
                    cookies = Uri.UnescapeDataString(Application.GetCookie(new Uri(url)));
                }
                catch (Exception)
                {
                    cookies = "";
                }

                string[] cs = cookies.Split(';');
                foreach (string c in cs)
                {
                    string[] cpr = c.Split('=');
                    if (cpr.Length == 2 && cpr[0].Trim() == "sid")
                    {
                        sid = cpr[1];
                    }
                }

                if (sid != "")
                {
                    Globals.Ribbons.Ribbon1.SFDebug("SessionId", sid);

                    string u = _url;
                    string p = u + "/services/Soap/u/" + _soapversion + "/" + _orgid;
                    string m = u + "/services/Soap/m/" + _soapversion + "/" + _orgid;
                    string s = sid;
                    bool? local = false;
                    string rtn = _d.Login(s, p, m, (local == true ? "access" : "sf"),_instancename);


                    if (rtn == "")
                    {
                        // get the settings
                        bool gotSettings = false;
                        DataReturn settings = _d.GetStaticResource("RibbonSettings");
                        if (!settings.success || settings.strRtn == "")
                        {
                            // get the default settings                      
                            var uri = new Uri("pack://application:,,,/AxiomIRISRibbon;component/Resources/Settings.json");
                            //string sjson = File.ReadAllText(uri.LocalPath);
                            byte[] bjson = AxiomIRISRibbon.Properties.Resources.Settings;
                            string sjson = Encoding.Default.GetString(bjson, 0, bjson.Length - 1);
                            try
                            {
                                _settings = new SForceEdit.Settings(sjson);
                                gotSettings = true;
                            }
                            catch (Exception eSet)
                            {
                                System.Windows.MessageBoxResult rslt = System.Windows.MessageBox.Show("Settings did not load : " + eSet.Message, "Axiom IRIS", System.Windows.MessageBoxButton.OK);
                            }
                        }
                        else
                        {
                            try
                            {
                                _settings = new SForceEdit.Settings(settings.strRtn);
                                gotSettings = true;
                            }
                            catch (Exception eSet)
                            {
                                System.Windows.MessageBoxResult rslt = System.Windows.MessageBox.Show("Settings did not load : " + eSet.Message, "Axiom IRIS", System.Windows.MessageBoxButton.OK);
                            }
                        }

                        if (!gotSettings) return;
                        Globals.ThisAddIn.setSettings(_settings);
                        Utility.HandleData(_d.LoadDefinitions());
                        Globals.Ribbons.Ribbon1.LoginOK();

                        // If this is not Prod then change the label on the Login
                        Globals.Ribbons.Ribbon1.btnLoginSSO.Label = _instance.ToString();
                        Globals.Ribbons.Ribbon1.sbtnLoginSSO.Label = _instance.ToString();

                        // update the settings to make this instance the default one
                        _local.Inst = _instance;
                        Globals.ThisAddIn.SaveLocalSettings(_local);

                        bsyInd.IsBusy = false;
                        this.Close();
                    }
                    else
                    {
                        bsyInd.IsBusy = false;
                        lblLoginMessage.Text = rtn;
                    }
                }
                else
                {
                    bsyInd.IsBusy = false;
                    lblLoginMessage.Text = "Problem getting the Session Id";
                }
            }

        }


        public void Login(LocalSettings.Instances ?Inst)
        {

            lblLoginMessage.Text = "";
            _webBrowser1.Navigate("about:blank");

            // set the busy indicator running
            bsyInd.IsIndeterminate = true;
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Logging In ...";

            _soapversion = _local.SoapVersion;

            if (!_local.Debug)
            {
                this.wfh1.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                this.wfh1.Visibility = System.Windows.Visibility.Visible;
            }
            
            if (Inst == null)
            {
                Inst = _local.Inst;                
            }

            _instance = Inst;
            
            if (Inst == LocalSettings.Instances.Dev)
            {                
                _url = _local.DevUrl;
                _orgid = _local.DevOrgId;
                _instancename = "SSO Dev (" + _url + ")";
            }
            else if (Inst == LocalSettings.Instances.IT)
            {              
                _url = _local.ITUrl;
                _orgid = _local.ITOrgId;
                _instancename = "SSO IT (" + _url + ")";
            }
            else if (Inst == LocalSettings.Instances.UAT)
            {             
                _url = _local.UATUrl;
                _orgid = _local.UATOrgId;
                _instancename = "SSO UAT (" + _url + ")";
            }
            else if (Inst == LocalSettings.Instances.Prod)
            {
                _url = _local.ProdUrl;
                _orgid = _local.ProdOrgId;
                _instancename = "SSO Prod (" + _url + ")";
            }


            // navigate the web page to the SSO
            string url = _url;

           // if not inside CS make this test.salesforce.com
           // url = "https://test.salesforce.com";            

            if (url != "")
            {
                try
                {
                    _webBrowser1.Navigate(url);                    
                    Utility.DoEvents();

                } catch(Exception e){

                    bsyInd.IsBusy = false;
                    lblLoginMessage.Text = "There has been a problem with the SSO login: " + e.Message;
                }
            }
            else
            {
                bsyInd.IsBusy = false;
                lblLoginMessage.Text = "No url setup for the SSO Login - please update settings";

            }
            
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Login(_instance);
        }








    }
}
