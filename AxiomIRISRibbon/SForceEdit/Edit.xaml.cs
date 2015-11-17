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
using System.Windows.Shapes;
using System.ComponentModel;
using Telerik.Windows.Controls;
using System.Data;
using System.Collections.ObjectModel;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for Edit.xaml
    /// </summary>
    public partial class Edit : RadWindow
    {
        SForceEdit.AxObject _axObj;

        public Edit(string sObject)
        {
            InitializeComponent();
            Utility.setTheme(this);
            _axObj = new SForceEdit.AxObject(sObject, this);
            g1.Children.Add(_axObj);


            this.Show();
            var window = this.ParentOfType<Window>();
            window.ShowInTaskbar = true;
            window.Title = "IRIS - Edit";
            var uri = new Uri("pack://application:,,,/AxiomIRISRibbon;component/Resources/Iris-Logo-Solo-Orange-40.png");
            window.Icon = BitmapFrame.Create(uri);
        }

        public Edit(string sObject,string Id)
        {
            InitializeComponent();
            Utility.setTheme(this);
            _axObj = new SForceEdit.AxObject(sObject, this,Id);
            g1.Children.Add(_axObj);

            this.Show();
            var window = this.ParentOfType<Window>();
            window.ShowInTaskbar = true;
            window.Title = "IRIS - Edit";
            var uri = new Uri("pack://application:,,,/AxiomIRISRibbon;component/Resources/Iris-Logo-Solo-Orange-40.png");
            window.Icon = BitmapFrame.Create(uri);
        }

        public Edit(string Mode,string sObject, string Id)
        {
            InitializeComponent();
            Utility.setTheme(this);
            if (Mode != "")
            {
                _axObj = new SForceEdit.AxObject(Mode,sObject, this, Id);
                g1.Children.Add(_axObj);
                this.Show();
                var window1 = this.ParentOfType<Window>();
                window1.ShowInTaskbar = true;
                window1.Title = "IRIS - Zoom";
                var uri = new Uri("pack://application:,,,/AxiomIRISRibbon;component/Resources/Iris-Logo-Solo-Orange-40.png");
                window1.Icon = BitmapFrame.Create(uri);
                window1.Show();
                window1.Activate();

                window.Show();
                window.Focus();
                window.Top = window.Top - 20;

                // having issues getting the window to come to the front - find this trick
                // on stachexchange basically just waits and then brings
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    window1.Activate();
                }), System.Windows.Threading.DispatcherPriority.ContextIdle, null);
            }
           
        }

        public void OpenZoomEditId(string Id)
        {
            this._axObj.LoadDataZoom(Id);            
            window.Show();
            window.Focus();
            window.Top = window.Top - 20;
        }


    }
}
