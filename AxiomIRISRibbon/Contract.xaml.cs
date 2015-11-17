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
    /// Interaction logic for Contract.xaml
    /// </summary>
    public partial class Contract : Window
    {

        private Data _d;

        public Contract()
        {
            InitializeComponent();
            Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();

            RefreshContractList();

            btnCancel.IsEnabled = false;
            btnSave.IsEnabled = false;
            btnOpen.IsEnabled = false;


            if (!Globals.ThisAddIn.getDebug()) tbHidden.Visibility = System.Windows.Visibility.Hidden;
        }


        public void Refresh()
        {
            this.RefreshContractList();
        }


        private void RefreshContractList()
        {
            DataReturn dr = Utility.HandleData(_d.GetVersions());
            if (!dr.success) return;
            dgContracts.ItemsSource = dr.dt.DefaultView;

        }


        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            //check the required fields
            if (tbName.Text == "")
            {
                MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                return;
            }

            DataRow drow;

            DataView dv = (DataView)dgContracts.ItemsSource;
            drow = dv.Table.NewRow();
            //Update from the form
            Utility.UpdateRow(new Grid[] { formGrid1, formGrid2 }, drow);

            //Save the values
            DataReturn dr = Utility.HandleData(_d.SaveContract(drow));
            if (!dr.success) return;
            tbId.Text = dr.id;


            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
            btnOpen.IsEnabled = true;

            RefreshContractList();
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
                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2 }, ((DataRowView)dgContracts.SelectedItem).Row);
            }
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }


        private void ClauseRowDoubleClick(object sender, RoutedEventArgs e)
        {
            Open(tbId.Text, tbTemplate__c.Text, tbTemplate__r_Name.Text, tbTemplate__r_PlaybookLink__c.Text);
            this.Hide();
        }

        private void dgContracts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgContracts.SelectedIndex > -1)
            {
                if (btnSave.IsEnabled)
                {
                    MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.Cancel)
                    {
                        dgContracts.SelectedIndex = -1;
                        return;
                    }
                }

                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2 }, ((DataRowView)dgContracts.SelectedItem).Row);
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                btnOpen.IsEnabled = true;
                tbXML.Text = "";
            }
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
            NewContract();
        }

        public void NewContract()
        {
            if (btnSave.IsEnabled)
            {
                MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    dgContracts.SelectedIndex = -1;
                    return;
                }
            }

            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2 });
            dgContracts.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;
            tbXML.Text = "";
        }

        private void btnReload_Click(object sender, RoutedEventArgs e)
        {
            RefreshContractList();
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult res = MessageBox.Show("Are you sure?", "Warning", MessageBoxButton.OKCancel);
            if (res == MessageBoxResult.Cancel)
            {
                return;
            }

            //Delete!
            _d.DeleteTemplate(tbId.Text);
            RefreshContractList();
            dgContracts.SelectedIndex = -1;
            btnOpen.IsEnabled = false;
            tbXML.Text = "";
        }

        public void Open(string Id, string TemplateId, string TemplateName, string TemplatePlaybookLink)
        {

            Globals.ThisAddIn.ProcessingStart("Create New Contract");

            //Get the template details 
            Globals.ThisAddIn.ProcessingUpdate("Get the Template File");
            this.Hide(); //close first or it will switch back to this doc

            string filename;
            if (Id == "")
            {
                filename = Utility.SaveTempFile(TemplateId);
                filename = Utility.HandleData(_d.GetTemplateFile(TemplateId, filename)).strRtn;
            }
            else
            {
                filename = Utility.SaveTempFile(Id);
                filename = Utility.HandleData(_d.GetDocumentFile(Id, filename)).strRtn;
            }

            //Open the doc
            Word.Document doc = null;
            try
            {
                doc = Globals.ThisAddIn.Application.Documents.Open(filename);
                Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ProcessingStop("Stop");
            }
            
            //If this is the template then change
            if (Id == "") Globals.ThisAddIn.AddDocId(doc, "Contract", "");

            Globals.ThisAddIn.AddContractContentControlHandler(doc);

            //if the action bar isn't open then open it
            Globals.ThisAddIn.ShowTaskPane(true);

            ContractEdit.SForceEditSideBar2 u = Globals.ThisAddIn.GetTaskPaneControlContract();
            if (u != null)
            {
                if (Id != "")
                {
                    // TODO - not really sure where this would be used!  might need to get the attachment id
                    u.BuildSideBarFromVersion(Id, "Attached", "");
                }
                else
                {
                    u.BuildSideNoVersion(TemplateId, TemplateName, TemplatePlaybookLink);
                }
            }

            //Scroll to the top
            Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = true;
            Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = 0;            
            Globals.ThisAddIn.ProcessingStop("Stop");

            if(doc!=null) doc.Activate();
        }


        public void Open(string Id, string TemplateId, string TemplateName, string TemplatePlaybookLink,string VersionId)
        {

            Globals.ThisAddIn.ProcessingStart("Create New Contract");

            //Get the template details 
            Globals.ThisAddIn.ProcessingUpdate("Get the Template File");
            this.Hide(); //close first or it will switch back to this doc

            string filename;
            if (Id == "")
            {
                filename = Utility.SaveTempFile(TemplateId);
                filename = Utility.HandleData(_d.GetTemplateFile(TemplateId, filename)).strRtn;
            }
            else
            {
                filename = Utility.SaveTempFile(Id);
                filename = Utility.HandleData(_d.GetDocumentFile(Id, filename)).strRtn;
            }

            //Open the doc
            Word.Document doc = null;
            try
            {
                doc = Globals.ThisAddIn.Application.Documents.Open(filename);
                Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ProcessingStop("Stop");
            }

            //If this is the template then change
            if (Id == "") Globals.ThisAddIn.AddDocId(doc, "Contract", "");

            Globals.ThisAddIn.AddContractContentControlHandler(doc);

            //if the action bar isn't open then open it
            Globals.ThisAddIn.ShowTaskPane(true);

            ContractEdit.SForceEditSideBar2 u = Globals.ThisAddIn.GetTaskPaneControlContract();
            if (u != null)
            {
                if (Id != "")
                {
                    // TODO - not really sure where this would be used!  might need to get the attachment id
                    u.BuildSideBarFromVersion(Id, "Attached", "");
                }
                else
                {
                    u.BuildSideNoVersion(TemplateId, TemplateName, TemplatePlaybookLink);
                }

             
            }

            //Scroll to the top
            Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = true;
            Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = 0;
            Globals.ThisAddIn.ProcessingStop("Stop");

            if (doc != null) doc.Activate();
        }

        public void OpenClauseFromNegotiatedDoc(string ContractId)
        {
            //What do we do - get the id, then look up the template - open the sidebar using the tempalte
            //then step through clauses and elements and get the values from the doc!


            Globals.ThisAddIn.ProcessingStart("Load Document");
            this.Hide(); //close first or it will switch back to this doc

            DataReturn dr = Utility.HandleData(_d.GetVersion(ContractId));
            string TemplateId = dr.dt.Rows[0]["Template__c"].ToString();
            string TemplateName = dr.dt.Rows[0]["Template__r_Name"].ToString();
            string TemplatePlaybookLink = dr.dt.Rows[0]["Template__r_PlaybookLink__c"].ToString();
            //Get the template details 
            Globals.ThisAddIn.ProcessingUpdate("Get the Template File");



            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            //Set the id to be a contract rather than an exported
            Globals.ThisAddIn.AddDocId(doc, "Contract", ContractId);

            //Globals.ThisAddIn.AddContractContentControlHandler(doc);

            //if the action bar isn't open then open it
            Globals.ThisAddIn.ShowTaskPane(true);

            ContractEdit.SForceEditSideBar2 u = Globals.ThisAddIn.GetTaskPaneControlContract();
            if (u != null)
            {
                u.BuildSideBar(TemplateId, TemplateName, TemplatePlaybookLink);

                //TODO load the values from the document
                u.LoadContractDataFromNegotiatedDoc(ContractId, dr.dt.Rows[0]["Name"].ToString());
            }

            //Scroll to the top
            Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = 0;
            Globals.ThisAddIn.ProcessingStop("Stop");

            doc.Activate();


            

            return;
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            Open(tbId.Text, tbTemplate__c.Text, tbTemplate__r_Name.Text, tbTemplate__r_PlaybookLink__c.Text);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }


        public void CreateNewVersion(string objname,string id, string templatename){



        }


    }
}
