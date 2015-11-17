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
using System.ComponentModel;

using System.Web;
using System.Net;
using System.IO;
using System.Web.Services;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class Login : RadWindow
    {
        private Data _d;
        private SForceEdit.Settings _settings;
        private LocalSettings _local;

        BackgroundWorker backgroundWorker;

        string[] loginentries;

        string passPhrase = "jkloekdlx23:dd";
        
        public Login()
        {
            InitializeComponent();
            Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();
            _settings = Globals.ThisAddIn.getSettings();
            _local = Globals.ThisAddIn.GetLocalSettings();

            // depending on the setting hide the demo logins
            GetLogins(_local.ShowAllLogins);

            if (radComboDemoLogins.Items.Count > 0)
            {
                radComboDemoLogins.SelectedIndex = 0;
            }

            // hid the theme pcicker and the local dbase
            themepick1.Visibility = System.Windows.Visibility.Hidden;
            label4.Visibility = System.Windows.Visibility.Hidden;
            themepick1.SelectedIndex = 0;
            cbLocal.Visibility = System.Windows.Visibility.Hidden;

            tbUserName.Focus();
            tbUserName.SelectAll();
        }


        private void GetLogins(bool get){
            //Get the Logins from the webpage
            radComboDemoLogins.Items.Clear();

            if (get)
            {
                try
                {
                    string url = "http://axiomtest.azurewebsites.net/RibbonLogginsV2.txt";
                    string logins = GetPage(url);

                    loginentries = logins.Split('\n');

                    foreach (string en in loginentries)
                    {
                        string[] v = en.Split('|');
                        radComboDemoLogins.Items.Add(v[0]);
                    }

                }
                catch (Exception)
                {

                }
            }
            else
            {
                // Hide and don' get
                radComboDemoLogins.Visibility = System.Windows.Visibility.Hidden;
                lblDemoLogins.Visibility = System.Windows.Visibility.Hidden;
                btnDemoLogins.Visibility = System.Windows.Visibility.Hidden;
                cbLocal.Visibility = System.Windows.Visibility.Hidden;
            }

            return;

        }


        private string GetPage(string url)
        {
            //Helper Function - runs the url with the cookies and all the default settings
            //For a simple Get

            try
            {
                string rawLocation = url;
                Uri location = new Uri(rawLocation);

                HttpWebRequest webRequest = WebRequest.Create(location) as HttpWebRequest;
                webRequest.Method = "GET";
                webRequest.ProtocolVersion = HttpVersion.Version11;
                webRequest.AllowAutoRedirect = true;
                webRequest.KeepAlive = true;
                webRequest.Referer = "";
                webRequest.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.7) Gecko/20040626 Firefox/0.8";

                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                // read response stream and dump to string
                Stream streamResponse = webResponse.GetResponseStream();
                StreamReader streamRead = new StreamReader(streamResponse);

                string output = streamRead.ReadToEnd();
                //Debug.WriteLine(output);
                return output;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }

        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (backgroundWorker != null && backgroundWorker.IsBusy) return;

            //Login
            string u = tbUserName.Text;
            string p = tbPassword.Password;
            string t = tbToken.Text;
            string url = tbEndPoint.Text;
            bool? local = cbLocal.IsChecked;
            string instance = radComboDemoLogins.Text;

            bsyInd.IsIndeterminate = true;
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Logging In ...";

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += (obj, ev) => WorkerDoWork(obj, ev, u, p, t,url, local,instance);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
            backgroundWorker.RunWorkerAsync();

        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            backgroundWorker.DoWork -= (obj, ev) => WorkerDoWork(obj, ev, "", "", "", "", false, "");
            backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;
            string rtn = (string)e.Result;

            if (rtn == "")
            {

                    
                    // get the settings
                    bool gotSettings = false;
                    DataReturn settings = _d.GetStaticResource("RibbonSettings");
                    if (!settings.success || settings.strRtn == "")
                    {
                        // get the default settings                      
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
                btnOK.IsEnabled = true;
                this.Close();
            }
            else
            {
                lblLoginMessage.Text = rtn;
            }

            btnOK.IsEnabled = true;

        }


        void WorkerDoWork(object sender, DoWorkEventArgs e, string UserName, string Password, string Token,string Url,bool? local,string InstanceDesc)
        {
            string rtn = _d.Login(UserName, Password, Token,Url,InstanceDesc);
            e.Result = rtn;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void RadComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string theme = themepick1.Text;
            
            if (theme == "Dark")
            {
                StyleManager.ApplicationTheme = new Expression_DarkTheme();
            }
            else if (theme == "Office")
            {
                StyleManager.ApplicationTheme = new Office_BlackTheme();
  
            }
            else
            {
                StyleManager.ApplicationTheme = new Windows8Theme();

            }
            Utility.setTheme(this);
        }

        private void radComboDemoLogins_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string demo = radComboDemoLogins.Text;
            if (loginentries != null)
            {
                foreach (string en in loginentries)
                {
                    string[] v = en.Split('|');
                    if (demo == v[0])
                    {

                        string password = v.Length >= 2 ? v[2].Trim() : "";
                        if (password != "") password = Utility.Decrypt(password, passPhrase);

                        tbUserName.Text = v.Length >= 1 ? v[1].Trim() : "";
                        tbPassword.Password = password;
                        tbToken.Text = v.Length >= 3 ? v[3].Trim() : "";
                        tbEndPoint.Text = v.Length >= 4 ? v[4].Trim() : "";
                    }
                }
            }
        }

        private void button1_Click_1(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("The login config line has been pasted to the clipboard with password encrypted:\n" + radComboDemoLogins.Text + "|" + tbUserName.Text + "|" + Utility.Encrypt(tbPassword.Password, passPhrase) + "|" + tbToken.Text + "|" + tbEndPoint.Text);
            Clipboard.SetText(radComboDemoLogins.Text + "|" + tbUserName.Text + "|" + Utility.Encrypt(tbPassword.Password, passPhrase) + "|" + tbToken.Text + "|" + tbEndPoint.Text);
        }
    }
}
