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


namespace AxiomIRISRibbon
{


    /// <summary>
    /// Interaction logic for Settings.xaml
    /// </summary>
    public partial class Settings : RadWindow
    {

        public Settings()
        {
            InitializeComponent();
            Utility.setTheme(this);
            
            LocalSettings s = Globals.ThisAddIn.GetLocalSettings();
            propertyGrid.Item = s;

            // tell the settings where this is being called from
            s.settingsDialog = this;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            LocalSettings ls = (LocalSettings)propertyGrid.Item;
            ls.Reset();

            propertyGrid.Item = ls;

        }

        public void SetTheme()
        {
            Globals.ThisAddIn.SetTheme();
            Utility.setTheme(this);
        }

        public void SetInstance(LocalSettings.Instances ?inst)
        {
            if (inst != null && inst != LocalSettings.Instances.Prod)
            {
                Globals.Ribbons.Ribbon1.btnLoginSSO.Label = inst.ToString();
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Label = inst.ToString();
            }
        }

        public void SetSSOLogin(bool value)
        {
            if (value)
            {
                Globals.Ribbons.Ribbon1.btnLogin.Visible = false;
                Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = true;
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = true;
            }
            else
            {
                Globals.Ribbons.Ribbon1.btnLogin.Visible = true;
                Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = false;
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = false;
            }
        }

        public void SetShowAllLogins(bool value)
        {
            if (value)
            {
                Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = false;
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = true;
            }
            else
            {
                Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = true;
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = false;
            }
        }

        public void SetDebug(bool value)
        {
            if (value)
            {
                Globals.Ribbons.Ribbon1.btnLogin.Visible = true;
                
                Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = true;
                Globals.Ribbons.Ribbon1.btnLoginSSO.Label = "SSO";

                Globals.Ribbons.Ribbon1.gpDebug.Visible = true;
            }
            else
            {
               
                Globals.Ribbons.Ribbon1.btnLoginSSO.Label = "Login";
                Globals.Ribbons.Ribbon1.gpDebug.Visible = false;


                LocalSettings ls = (LocalSettings)propertyGrid.Item;
                if (ls.SSOLogin)
                {
                    Globals.Ribbons.Ribbon1.btnLogin.Visible = false;
                    Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = true;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.btnLogin.Visible = true;
                    Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = false;
                }


            }

        }
    }
}
