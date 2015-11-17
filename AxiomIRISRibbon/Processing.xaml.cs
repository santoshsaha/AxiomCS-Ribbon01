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

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for Processing.xaml
    /// </summary>
    public partial class Processing : Window
    {
        int _i;

        public Processing()
        {
            InitializeComponent();
        }

        
        public void Update(string t)
        {
            tbStatus.Text = tbStatus.Text + "\n" + _i.ToString() + ". \t" + t + "...";
            _i++;
            sv1.ScrollToBottom();
            Utility.DoEvents();
        }

        public void Start(string t){
            _i = 1;
            tbStatus.Text = t;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            this.Show();
        }

        public void Stop(string t)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            this.Hide();
            Utility.DoEvents();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }
    }
}
