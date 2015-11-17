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
    /// Interaction logic for AboutReleaseNotes.xaml
    /// </summary>
    public partial class AboutReleaseNotes : RadWindow
    {
        public AboutReleaseNotes()
        {
            InitializeComponent();
            
            Utility.setTheme(this);

            rtNotes.Text = AxiomIRISRibbon.Properties.Resources.ReleaseNotes;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
