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
    public partial class Template : Window
    {
        private string editmode;
        private Data d;

        public Template()
        {
            InitializeComponent();
            Utility.setTheme(this);


            d = Globals.ThisAddIn.getData();

            //Get the drop down values and populate the list box
            GetDropDown();
            RefreshTemplateList();
            btnCancel.IsEnabled = false;
            btnSave.IsEnabled = false;
            btnOpen.IsEnabled = false;

            editmode = "edit";

            Utility.ReadOnlyForm(true, new Grid[] { formGrid1, formGrid2 });


            if (!Globals.ThisAddIn.getDebug()) tbHidden.Visibility = System.Windows.Visibility.Hidden;
        }

        public void Refresh()
        {
            RefreshTemplateList();
        }


        private void GetDropDown()
        {
            DataTable dt = Utility.HandleData(d.GetPickListValues("RibbonTemplate__c", "Type__c", false)).dt;

            ComboBoxItem i = new ComboBoxItem();
            i.Content = "";
            cbType.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new ComboBoxItem();
                i.Content = r["Value"].ToString();
                cbType.Items.Add(i);
            }

            dt = Utility.HandleData(d.GetPickListValues("RibbonTemplate__c", "State__c", false)).dt;

            i = new ComboBoxItem();
            i.Content = "";
            cbState.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new ComboBoxItem();
                i.Content = r["Value"].ToString();
                cbState.Items.Add(i);
            }

        }

        private void LoadTemplatesDLL()
        {
            try
            {
                d = Globals.ThisAddIn.getData();
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(d.GetTemplates(true));
                if (!dr.success) return;

                DataTable dt1 = dr.dt;
                //cbAmendmewnt.Items.Clear();

                ComboBoxItem i;

                // RadComboBoxItem selected = null;
                foreach (DataRow r in dt1.Rows)
                {
                    i = new ComboBoxItem();
                    i.Tag = r["Id"].ToString();
                    i.Content = r["Name"].ToString();
                    this.cbAmendmewnt.Items.Add(i);

                }

            }
            catch (Exception ex)
            {

            }
        }

        public void UpdateOptions(ComboBox c, string[] options)
        {
            c.Items.Clear();
            foreach (string option in options)
            {
                ComboBoxItem i = new ComboBoxItem();
                i.Content = option;
                c.Items.Add(i);
            }
        }

        private void RefreshTemplateList()
        {
            DataReturn dr = Utility.HandleData(d.GetTemplates(false));
            if (!dr.success) return;
            dgTemplates.ItemsSource = dr.dt.DefaultView;

        }


        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            //check the required fields
            if (tbName.Text == "")
            {
                MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                return;
            }

            Globals.ThisAddIn.ProcessingStart("Save ...");

            DataRow drow;

            DataView dv = (DataView)dgTemplates.ItemsSource;
            drow = dv.Table.NewRow();
            //Update from the form
            Utility.UpdateRow(new Grid[] { formGrid1, formGrid2 }, drow);

            //Save the values
            DataReturn dr = Utility.HandleData(d.SaveTemplate(drow));
            if (!dr.success) return;
            tbId.Text = dr.id;

            Globals.ThisAddIn.ProcessingStart("Save Template Data");

            //Take the active doc and save it with the new id
            if (editmode == "newfromdoc")
            {
                // close the dialog
                this.Hide();

                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Globals.ThisAddIn.AddDocId(doc, "ContractTemplate", tbId.Text);

                try
                {
                    Word.Style s = doc.Styles.Add("ContentControl");
                    s.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightOrange;
                }
                catch (Exception)
                {

                }

                //save the file - this will throw an error because the event handler will do the save
                //should be a way to do this without the throw
                try
                {
                    doc.SaveAs2("Dummy", Word.WdSaveFormat.wdFormatXMLDocument);
                }
                catch (Exception)
                {
                }

                Globals.ThisAddIn.ProcessingStart("Show Task Pane");

                //now open the sidebar and add the handler
                Globals.ThisAddIn.AddContentControlHandler(doc);
                Globals.ThisAddIn.ShowTaskPane(true);

                //reload the tree
                TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate();
                tsb.Refresh();

                Globals.ThisAddIn.ProcessingStop("");

            }

            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
            btnOpen.IsEnabled = true;

            Globals.ThisAddIn.ProcessingStart("Refresh Lists");

            RefreshTemplateList();
            Globals.Ribbons.Ribbon1.RefreshTemplatesList();

            Globals.ThisAddIn.ProcessingStop("Finished");
        }

        public void NewTemplate()
        {
            editmode = "newfromdoc";
            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
            Utility.ReadOnlyForm(false, new Grid[] { formGrid1, formGrid2 });
            dgTemplates.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;
        }



        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //If there is an id then reload from the grid - if not then just clear
            if (tbId.Text == "")
            {
                Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
                Utility.ReadOnlyForm(true, new Grid[] { formGrid1, formGrid2 });
            }
            else
            {
                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2 }, ((DataRowView)dgTemplates.SelectedItem).Row);
            }
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }



        private void ClauseRowDoubleClick(object sender, RoutedEventArgs e)
        {
            Open();
        }

        private void dgTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Utility.ReadOnlyForm(false, new Grid[] { formGrid1, formGrid2 });

            if (dgTemplates.SelectedIndex > -1)
            {
                if (btnSave.IsEnabled)
                {
                    MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.Cancel)
                    {
                        dgTemplates.SelectedIndex = -1;
                        return;
                    }
                }

                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2 }, ((DataRowView)dgTemplates.SelectedItem).Row);

                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                btnOpen.IsEnabled = true;
                tbXML.Text = "";
            }
            else
            {
                Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                btnOpen.IsEnabled = false;
                tbXML.Text = "";
            }

            LoadTemplatesDLL();
        }


        private void formTextBoxChanged(object sender, TextChangedEventArgs e)
        {
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
                    dgTemplates.SelectedIndex = -1;
                    return;
                }
            }

            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
            dgTemplates.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;
            tbXML.Text = "";
        }

        private void btnReload_Click(object sender, RoutedEventArgs e)
        {
            RefreshTemplateList();
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult res = MessageBox.Show("Are you sure?", "Warning", MessageBoxButton.OKCancel);
            if (res == MessageBoxResult.Cancel)
            {
                return;
            }

            //Delete!
            d.DeleteTemplate(tbId.Text);
            RefreshTemplateList();
            dgTemplates.SelectedIndex = -1;
            btnOpen.IsEnabled = false;
            tbXML.Text = "";
        }



        private void Open()
        {
            //Create new doc and insert the xml
            //string xml = HandleData(d.GetTemplateXML(tbId.Text)).strRtn;

            //TODO - have a scratch document for use of the toolbar
            //to stop having millions of docs open - need to add the management
            //Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //OK - using the docx and files because they are both smaller *and* 
            //they can then be viewed in native salesforce


            Globals.ThisAddIn.ProcessingStart("Opening Contract Template");
            this.Hide(); //close first or it will switch back to this doc

            string filename = Utility.SaveTempFile(tbId.Text);
            Globals.ThisAddIn.ProcessingUpdate("Download Template File From SForce");


            DataReturn dr = Utility.HandleData(d.GetTemplateFile(tbId.Text, filename));
            if (!dr.success) return;
            filename = dr.strRtn;

            Globals.ThisAddIn.ProcessingUpdate("Got Template File From SForce");

            Word.Document doc = null;
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

                try
                {
                    doc = Globals.ThisAddIn.Application.Documents.Open(filename);
                    Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problem Opening Template:" + ex.Message);
                    Globals.ThisAddIn.ProcessingStop("Finished");
                    return;
                }
            }

            Globals.ThisAddIn.AddContentControlHandler(doc);
            Globals.ThisAddIn.ShowTaskPane(true);

            //reload the tree
            TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate();
            tsb.Refresh();

            Globals.ThisAddIn.ProcessingStop("Finished");
        }


        public void OpenImportTemplate()
        {

            Globals.ThisAddIn.ProcessingStart("Import Contract Template");
            this.Hide(); //close first or it will switch back to this doc

            //reload the tree
            Globals.ThisAddIn.AddContentControlHandler(Globals.ThisAddIn.Application.ActiveDocument);
            Globals.ThisAddIn.ShowTaskPane(true);
            TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate();
            tsb.Import();

        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            Open();
        }

        //Stop the window being closed - jsut hide
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

    }
}
