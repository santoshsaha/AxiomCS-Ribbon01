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
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Data;

namespace AxiomIRISRibbon.TemplateEdit
{
    /// <summary>
    /// Interaction logic for CloneTemplate.xaml
    /// </summary>
    public partial class CloneTemplate : Telerik.Windows.Controls.RadWindow
    {

        private TEditSidebar sidebar;

        public CloneTemplate()
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);
        }

        public void Open(TEditSidebar ts){
            this.sidebar = ts;
            this.Clone2.IsChecked = true;
            this.tbCloneName.Text = ts.tbTemplateName.Text + "-Copy";
            this.tbCloneName.IsEnabled = true;
            this.txtPrepend.Text = "Copy";
        }

        private void Clone1_Checked(object sender, RoutedEventArgs e)
        {
            this.txtPrepend.IsEnabled = false;
        }

        private void Clone2_Checked(object sender, RoutedEventArgs e)
        {
            txtPrepend.IsEnabled = true;
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            string mode = "";
            if (this.Clone1.IsChecked==true)
            {
                mode = "CloneTemplate";
            }
            else
            {
                mode = "CloneTemplateConceptClause";
            }

            this.sidebar.DoClone(mode,this.tbCloneName.Text, this.txtPrepend.Text);
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }



    }
}
