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
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Data;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for Template.xaml
    /// </summary>
    public partial class Concept : Window
    {
        public string _editmode;
        private Data d;

        public Concept()
        {
            InitializeComponent();
            Utility.setTheme(this);

            d = Globals.ThisAddIn.getData();

            //populate the list box
            RefreshConceptList();
            btnCancel.IsEnabled = false;
            btnSave.IsEnabled = false;

            if (!Globals.ThisAddIn.getDebug()) tbHidden.Visibility = System.Windows.Visibility.Hidden;

            _editmode = "";
        }

        public void Refresh()
        {
            this.RefreshConceptList();
        }



        private void RefreshConceptList()
        {            
            DataReturn dr = Utility.HandleData(d.GetConcepts());
            if (!dr.success) return;
            dgConcepts.ItemsSource = dr.dt.DefaultView;

        }


        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            //check the required fields
            if (tbName.Text == "")
            {
                MessageBox.Show("Name field is required","Problem",MessageBoxButton.OK);
                return;
            }

            DataRow drow;

            DataView dv = (DataView)dgConcepts.ItemsSource;
            drow= dv.Table.NewRow();
            //Update from the form
            Utility.UpdateRow(new Grid[] { formGrid1, formGrid2 }, drow);
            
            //Save the values
            DataReturn dr = Utility.HandleData(d.SaveConcept(drow));
            if (!dr.success) return;
            tbId.Text = dr.id;

            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;

            RefreshConceptList();

            if (_editmode == "fromClause")
            {
                this.Hide();
                Clause cls = Globals.ThisAddIn.OpenClause(true,false);
                cls.RefreshConceptList(dr.id);
                
            }
        }

        public void NewConcept()
        {
            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
            dgConcepts.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
        }



        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
           //If there is an id then reload from the grid - if not then just clear
            if (tbId.Text == "")
            {
                Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
            }
            else
            {
                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2 }, ((DataRowView)dgConcepts.SelectedItem).Row);
            }
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }


        private void ClauseRowDoubleClick(object sender, RoutedEventArgs e)
        {

            this.Hide(); 
        }

        private void dgConcepts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgConcepts.SelectedIndex > -1)
            {
                if (btnSave.IsEnabled)
                {
                    MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.Cancel)
                    {
                        dgConcepts.SelectedIndex = -1;
                        return;
                    }
                }

                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2 }, ((DataRowView)dgConcepts.SelectedItem).Row);
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                tbXML.Text = "";
            }
        }


        private void formTextBoxChanged(object sender, TextChangedEventArgs e){
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
        }

        private void formComboChanged(object sender, SelectionChangedEventArgs e)
        {
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (btnSave.IsEnabled)
            {
                MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    dgConcepts.SelectedIndex = -1;
                    return;
                }
            }

            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
            dgConcepts.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            tbXML.Text = "";
        }

        private void btnReload_Click(object sender, RoutedEventArgs e)
        {
            RefreshConceptList();
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult res = MessageBox.Show("Are you sure?", "Warning", MessageBoxButton.OKCancel);
            if (res == MessageBoxResult.Cancel)
            {
                return;
            }

            //Delete!
            d.DeleteConcept(tbId.Text);
            RefreshConceptList();
            dgConcepts.SelectedIndex = -1;
            tbXML.Text = "";
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

            Globals.ThisAddIn.ProcessingStart("Opening Contract Template");
            this.Hide(); //close first or it will switch back to this doc

            string filename = Utility.SaveTempFile(tbId.Text);
            Globals.ThisAddIn.ProcessingUpdate("Download Template File From SForce");
            filename = Utility.HandleData(d.GetTemplateFile(tbId.Text, filename)).strRtn;
            Globals.ThisAddIn.ProcessingUpdate("Got Template File From SForce");

            Word.Document doc;
            if (filename == "")
            {
                doc = Globals.ThisAddIn.Application.Documents.Add();
                Word.Style s = doc.Styles.Add("ContentControl");
                s.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightOrange;

                //Add a document property so we know that is a contract template and what the id is                
                Globals.ThisAddIn.AddDocId(doc, "ContractTemplate", tbId.Text);
                Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }
            else
            {
                doc = Globals.ThisAddIn.Application.Documents.Open(filename);
                Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }

            Globals.ThisAddIn.AddContentControlHandler(doc);
            Globals.ThisAddIn.ShowTaskPane(true);

            //reload the tree
            TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate();
            tsb.Refresh();

            Globals.ThisAddIn.ProcessingStop("Finished");
        }

        //Stop the window being closed - jsut hide
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
            if (_editmode == "fromClause")
            {
                Clause cls = Globals.ThisAddIn.OpenClause(true,false);
                cls.RefreshConceptList("");
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            if (_editmode == "fromClause")
            {
                Clause cls = Globals.ThisAddIn.OpenClause(true, false);
                cls.RefreshConceptList("");
            }
        }

        private void cbAllowNone__c_Checked(object sender, RoutedEventArgs e)
        {
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
        }

        private void cbAllowNone__c_Unchecked(object sender, RoutedEventArgs e)
        {
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
        }



    }
}
