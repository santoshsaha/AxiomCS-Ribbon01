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
using System.Configuration;
using System.Deployment;
using Microsoft.Win32;
using System.Diagnostics;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class About : RadWindow
    {
        public About()
        {
            InitializeComponent();
            Utility.setTheme(this);

            string version = GetRunningVersion().ToString();
            this.tbVersion.Text = version.ToString();

            this.tbSF.Text = Globals.ThisAddIn.getData().GetInstanceInfo();
            this.tbUser.Text = Globals.ThisAddIn.getData().GetUserInfo();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private string GetRunningVersion()
        {

            string v = "IRIS Ribbon | Version - UNKOWN!";
            try
            {
                    System.Deployment.Application.ApplicationDeployment ad = System.Deployment.Application.ApplicationDeployment.CurrentDeployment;
                    Version vrn = ad.CurrentVersion;
                    v = "IRIS Ribbon | Version " + vrn.Major + "." + vrn.Minor + "." + vrn.Build + "." + vrn.Revision;
                
            }
            catch (Exception)
            {

            }

            return v;
        }

        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.OpenAboutReleaseNotes();                    
        }

        private void windowAbout_Activated(object sender, EventArgs e)
        {
            this.tbSF.Text = Globals.ThisAddIn.getData().GetInstanceInfo();
            this.tbUser.Text = Globals.ThisAddIn.getData().GetUserInfo();
        }

    }
}
