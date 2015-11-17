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
    /// Interaction logic for Clause.xaml
    /// </summary>
    public partial class Clause : Window
    {
        private Data _d;

        private string _templateid;
        private string _templatename;
        private string _templateclausemode;

        public Clause(bool refresh)
        {           
            InitializeComponent();
            Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();

            // Get the drop down values and populate the list box
            // Russel 1 Oct - this takes ages! don't need to do it if we don;t need the list, i.e. when we open a clause 
            // from a template so don't always do it
            if (refresh)
            {
                this.Refresh();
            }
            else
            {
                // no refresh means an Add or Clone and we always need the Concept list for that
                if (cbConcept.Items.Count == 0)
                {
                    RefreshConceptList("");
                }

                // need to set up the table if its not been done before
                if (dgClauses.ItemsSource == null)
                {
                    DataReturn dr = Utility.HandleData(_d.GetClause("123456789012345"));
                    if (!dr.success) return;
                    dgClauses.ItemsSource = dr.dt.DefaultView;
                }

            }

            


            if (!Globals.ThisAddIn.getDebug()) tbHidden.Visibility = System.Windows.Visibility.Hidden;

            btnCancel.IsEnabled = false;
            btnSave.IsEnabled = false;
            btnOpen.IsEnabled = false;
        }


        public void RefreshIfNotLoaded()
        {
            if (dgClauses.Items.Count == 0)
            {
                this.Refresh();        
            }
        }

        public void Refresh(){
            Globals.ThisAddIn.ProcessingStart("Load Concepts and Clauses");
            Globals.ThisAddIn.ProcessingUpdate("Get Concepts");
            this.RefreshConceptList("");
            Globals.ThisAddIn.ProcessingUpdate("Get Clauses");
            this.RefreshClauseList();
            Globals.ThisAddIn.ProcessingStop("Done");

            
        }


        public void SetFromTemplate(string templateid,string templatename,string mode){
            _templateid = templateid;
            _templatename = templatename;
            _templateclausemode = mode;
        }

        public void RefreshConceptList(string id)
        {
            DataReturn dr = Utility.HandleData(_d.GetConcepts());
            if (!dr.success) return;

            DataTable dt = dr.dt;
            cbConcept.Items.Clear();

            ComboBoxItem i = new ComboBoxItem();
            i.Content = "";
            cbConcept.Items.Add(i);

            ComboBoxItem selected = null;
            foreach (DataRow r in dt.Rows)
            {
                i = new ComboBoxItem();
                i.Tag = r["Id"].ToString();
                i.Content = r["Name"].ToString();
                cbConcept.Items.Add(i);

                if (r["Id"].ToString() == id) selected = i;
            }

            if (selected != null)
            {
                cbConcept.SelectedItem = selected;
            }
            else
            {
                cbConcept.SelectedIndex = -1;
            }
        }


        private void RefreshClauseList()
        {
            DataReturn dr = Utility.HandleData(_d.GetClauses());
            if (!dr.success) return;
            dgClauses.ItemsSource = dr.dt.DefaultView;
        }

        

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (btnSave.IsEnabled)
            {
                MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    dgClauses.SelectedIndex = -1;
                    return;
                }
            }

            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
            dgClauses.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;
            tbXML.Text = "";
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            this.Hide(); //close first or it will switch back to this doc

            //Will switch to having a clause as an object that we can pass around
            //at one point soon!
            if (tbId.Text != "")
            {
                OpenClause(tbId.Text);
            }

        }


        //Open the clause as a new document
        public void OpenClause(string Id)
        {

            string filename = Utility.SaveTempFile(Id);

            DataReturn dr = Utility.HandleData(_d.GetTemplateFile(Id, filename));
            if (!dr.success) return;
            filename = dr.strRtn;

            Word.Document doc;
            if (filename == "")
            {
                doc = Globals.ThisAddIn.Application.Documents.Add();
                Word.Style s = doc.Styles.Add("ContentControl");
                s.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightOrange;
                Globals.ThisAddIn.AddDocId(doc, "ClauseTemplate", Id);
            }
            else
            {
                doc = Globals.ThisAddIn.Application.Documents.Open(filename);
                Globals.ThisAddIn.AddDocId(doc, "ClauseTemplate", Id);
            }

            Globals.ThisAddIn.AddContentControlHandler(doc);

            Globals.ThisAddIn.ShowTaskPane(true);

            //reload the tree
            TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate();
            tsb.Refresh();
            
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //If there is an id then reload from the grid - if not then just clear
            if (tbId.Text == "")
            {
                Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
            }
            else
            {
                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2, formGrid3 }, ((DataRowView)dgClauses.SelectedItem).Row);
            }
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {

            Save();

            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
            btnOpen.IsEnabled = true;

            RefreshClauseList();
        }

        private string Save(){
            //check the required fields
            if (tbName.Text == "")
            {
                MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                return "";
            }
            if (cbConcept.Text == "")
            {
                MessageBox.Show("Concept field is required", "Problem", MessageBoxButton.OK);
                return "";
            }

            Globals.ThisAddIn.ProcessingStart("Save Clause");

            //Create a row from the list box to save the values
            DataRow drow;
            DataView dv = (DataView)dgClauses.ItemsSource;
            drow = dv.Table.NewRow();
            //Update from the form
            Utility.UpdateRow(new Grid[] { formGrid1, formGrid2, formGrid3 }, drow);

            //Save the values
            DataReturn dr = Utility.HandleData(_d.SaveClause(drow));
            if (!dr.success) return "";
            tbId.Text = dr.id;

           
            if (tbIdToCopy.Text != "")
            {
                string oldclauseid = tbIdToCopy.Text;
                string newclauseid = tbId.Text;
                string xml = "";
                //Saving a copy, have to get the XML and populate that
                //and copy all the elements
                Globals.ThisAddIn.ProcessingUpdate("Get the Clause Template File from SF for the Clause to Copy");
                string filename = Utility.SaveTempFile(newclauseid);
                dr = Utility.HandleData(_d.GetClauseFile(oldclauseid, filename));
                if (!dr.success) return "";
                filename = dr.strRtn;
                if (filename == "")
                {
                    xml = "";
                }
                else
                {
                    //This is the bit that causes the flash - have to open and close the file
                    Word.Document doc1 = Globals.ThisAddIn.Application.Documents.Open(filename, Visible: false);
                    xml = doc1.WordOpenXML;
                    var docclose = (Microsoft.Office.Interop.Word._Document)doc1;
                    docclose.Close();
                }

                tbXML.Text = xml;

                //Copy any elements
                Globals.ThisAddIn.ProcessingUpdate("Get the Elements for the old clause to copy");
                DataReturn dr1 = Utility.HandleData(_d.GetElements(oldclauseid));
                if (dr1.success)
                {
                    DataTable dtElement = dr1.dt;

                    if (dtElement.Rows.Count > 0)
                    {
                        foreach (DataRow er in dtElement.Rows)
                        {                            
                            string name = Utility.Truncate(tbName.Text, 35) + "-" + Utility.Truncate(er["Element__r_Name"].ToString(),35);                            
                            dr = Utility.HandleData(_d.SaveClauseElement("", name, newclauseid, er["Element__r_Id"].ToString(), er["Order__c"].ToString()));
                            if (!dr.success) return "";
                        }
                    }
                }

            }

            if (tbXML.Text != "")
            {
                //this is a new one - need to save the clause
                //remeber to add in the clause id as a property
                
                Word.Document doc = Globals.ThisAddIn.Application.Documents.Add();
                Globals.ThisAddIn.AddDocId(doc, "ClauseTemplate", tbId.Text);
                try
                {
                    doc.Range(0).InsertXML(tbXML.Text);
                }
                catch (Exception)
                {
                } 

                try
                {
                    if (!Utility.StyleExists(doc.Styles, "ContentControl"))
                    {
                        Word.Style s = doc.Styles.Add("ContentControl");
                        s.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightOrange;
                    }
                }
                catch (Exception)
                {
                }

                //save the file - this will throw an error because the event handler will do the save
                //should be a way to do this without the throw
                try
                {
                    doc.SaveAs2("Dummy", Word.WdSaveFormat.wdFormatXMLDocument);
                } catch(Exception){
                }
               
                var docclose = (Microsoft.Office.Interop.Word._Document)doc;                
                docclose.Close();
                  
            }

            Globals.ThisAddIn.ProcessingStop("Done");
            return tbId.Text;

            
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult res = MessageBox.Show("Are you sure?", "Warning", MessageBoxButton.OKCancel);
            if (res == MessageBoxResult.Cancel)
            {
                return;
            }

            //Delete!
            DataReturn dr = Utility.HandleData(_d.DeleteClause(tbId.Text));
            if (!dr.success) return;

            RefreshClauseList();
            dgClauses.SelectedIndex = -1;
            btnOpen.IsEnabled = false;
            tbXML.Text = "";
        }


        private void btnReload_Click(object sender, RoutedEventArgs e)
        {
            this.Refresh();
        }

        private void ClauseRowDoubleClick(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void dgClauses_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgClauses.SelectedIndex > -1)
            {
                if (btnSave.IsEnabled)
                {
                    MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.Cancel)
                    {
                        dgClauses.SelectedIndex = -1;
                        return;
                    }
                }

                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2, formGrid3 }, ((DataRowView)dgClauses.SelectedItem).Row);
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                btnOpen.IsEnabled = true;
                tbXML.Text = "";
            }
            else
            {
                Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                btnOpen.IsEnabled = false;
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


        public void NewClause(string Name,string Desc,string Concept,string Text,string Xml)
        {
            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
            dgClauses.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;

            //set the form values
            tbName.Text = Name;
            tbDescription.Text = Desc;
            tbText.Text = Text;
            tbXML.Text = Xml;

            //Set the Concept - create it if it doesn't exist
            if (Concept != "")
            {
                bool found = false;
                foreach (ComboBoxItem i in cbConcept.Items)
                {
                    if (i.Content.ToString().Equals(Concept,StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        cbConcept.Text = Concept;
                        break;
                    }
                }
                if (!found)
                {
                    _d.SaveConcept("", Concept, "", "","",false);
                    RefreshConceptList("");
                    cbConcept.Text = Concept;
                }
            }
        }

        private void btnInsert_Click(object sender, RoutedEventArgs e)
        {

            if (Globals.ThisAddIn.isTemplate())
            {

                Globals.ThisAddIn.ProcessingStart("Insert Clause");

                //save - exit if we don't get an id back
                string clauseid = tbId.Text;
                if (btnSave.IsEnabled)
                {
                    Globals.ThisAddIn.ProcessingUpdate("Save Clause");
                    clauseid = Save();
                    if (clauseid == "")
                    {
                        return;
                    }
                }

                string conceptid = ((ComboBoxItem)cbConcept.SelectedItem).Tag.ToString();
                string templateclauseid = "";

                //if this is an Insert then we will have the templateid
                if (_templateid != "")
                {
                    Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                    if (_templateclausemode == "clone")
                    {
                        // if we are cloning then we know that the control is there and we just have to save the link and update the tree
                        Globals.ThisAddIn.ProcessingUpdate("Add the Clause to the Template");

                        DataReturn dr = new DataReturn();

                        // Check if it is already there - this can happen if the control is removed
                        // from the document or if they try and add the clause twice
                        dr = Utility.HandleData(_d.GetTemplateClause(_templateid, clauseid));
                        if (dr.dt.Rows.Count == 0)
                        {
                            string name = Utility.Truncate(_templatename, 35) + "-" + Utility.Truncate(tbName.Text, 35);
                            dr = Utility.HandleData(_d.SaveTemplateClause("", name, _templateid, clauseid, "",""));
                            if (!dr.success) return;
                            templateclauseid = dr.id;
                        }
                        else
                        {
                            templateclauseid = dr.dt.Rows[0]["Id"].ToString();
                        }

                    }
                    else
                    {

                        // First things first - see if we can insert teh content control - if we can't then don't do anything else
                        // and flag it as a problem

                        Globals.ThisAddIn.ProcessingUpdate("Check the selection");

                        
                        Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                        Word.ContentControl c;

                        try
                        {
                            
                            // ok we get an error if there is a pagebreak followed directly by the section we are trying to add
                            // so add a paragraph
                            sel.Range.InsertBefore("\r");

                            c = sel.Document.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText);
                        }
                        catch (Exception ex)
                        {
                            // get rid of the return we just inserted
                            try
                            {
                                if (sel.Range.Start > 0) doc.Range(sel.Range.Start - 1, sel.Range.Start).Delete();
                            }
                            catch (Exception ex2)
                            {

                            }
                            Globals.ThisAddIn.ProcessingUpdate("Cannot create a clause holder in the current selection");
                            Globals.ThisAddIn.ProcessingStop("End");
                            MessageBox.Show("Sorry, can't create a clause holder with the current selection! :" + ex.Message);
                            return;
                        }

                        // get rid of the return we just inserted - does somehting odd when its the first letter
                        if (sel.Range.Start > 0) doc.Range(sel.Range.Start - 1, sel.Range.Start).Delete();

                        //Add to Contract Template
                        Globals.ThisAddIn.ProcessingUpdate("Add the Clause to the Template");

                        DataReturn dr = new DataReturn();

                        // Check if it is already there - this can happen if the control is removed
                        // from the document or if they try and add the clause twice
                        dr = Utility.HandleData(_d.GetTemplateClause(_templateid, clauseid));
                        if (dr.dt.Rows.Count == 0)
                        {
                            string name = Utility.Truncate(_templatename, 35) + "-" + Utility.Truncate(tbName.Text, 35);
                            dr = Utility.HandleData(_d.SaveTemplateClause("", name, _templateid, clauseid, "",""));
                            if (!dr.success) return;
                            templateclauseid = dr.id;
                        }
                        else
                        {
                            templateclauseid = dr.dt.Rows[0]["Id"].ToString();
                        }

                        //Check we don't already have this concept in the template - we just added one so make sure we don't have 2
                        Globals.ThisAddIn.ProcessingUpdate("Check if we have this concept");
                        if (Utility.HandleData(_d.GetTemplateClauseCount(_templateid, conceptid)).dt.Rows.Count > 1)
                        {
                            if (_templateclausemode == "inplace")
                            {
                                //Just remove the selection - it is now incorporated in the concept 
                                Globals.ThisAddIn.ProcessingUpdate("Remove Selection");
                                sel.Delete();
                                try
                                {
                                    c.Delete();
                                }
                                catch (Exception)
                                {

                                }
                            }
                        }
                        else
                        {
                            // Add in the new concept  
                            // fix can only be 64 chars!
                            c.Title = Utility.Truncate(cbConcept.Text,64);
                            c.Tag = "Concept|" + conceptid.ToString();
                            c.LockContentControl = true;
                            c.LockContents = true;
                        }
                    }
                    
                   
                    //Save! important to try and keep things in sync
                    try
                    {
                        doc.Save();
                    }
                    catch (Exception)
                    {
                    }

                    
                }

                // Update the list bars - This reloads everything from Salesforce for all open templates
                // so too **slow** - only update mentions of *this* clause
                // Globals.ThisAddIn.ProcessingStop("Refresh All Open Contracts");
                // Globals.ThisAddIn.RefreshAllTaskPanes();

                // For now just relaod this one
                // this is also too slow!
                // need to just do this **CLAUSE**

                // update the concept list as we may have created a new one
                Globals.ThisAddIn.GetTaskPaneControlTemplate().RefreshConceptList();

                Globals.ThisAddIn.ProcessingUpdate("Refresh Tree");
                Globals.ThisAddIn.GetTaskPaneControlTemplate().RefreshClause(clauseid,templateclauseid);
                
                this.Hide();

                Globals.ThisAddIn.ProcessingStop("Done");
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }

        private void btnAddConcept_Click(object sender, RoutedEventArgs e)
        {
            Concept c = Globals.ThisAddIn.OpenConcept();
            c.NewConcept();
            c._editmode = "fromClause";

        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            //Set the ID to blank and add Copy to the Name
            tbIdToCopy.Text = tbId.Text;
            tbId.Text = "";
            tbName.Text += "-Copy";

            dgClauses.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;
        }


        public void CloneClause(string Name, string Desc, string ConceptId, string IdToCopy)
        {
            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
            dgClauses.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();
            btnOpen.IsEnabled = false;

            //set the form values
            tbIdToCopy.Text = IdToCopy;
            tbName.Text = Name + "-Copy";
            tbDescription.Text = Desc;

            // Russel Oct 14 - this was using name to pick the right clause
            // caused issues if there were duplicate clause names - change to use
            // id

            //Set the Concept - create it if it doesn't exist
            if (ConceptId != "")
            {
                bool found = false;
                ComboBoxItem selecteditem = null;
                foreach (ComboBoxItem i in cbConcept.Items)
                {
                    if (i.Tag != null)
                    {
                        string cbid = i.Tag.ToString();
                        if (cbid == ConceptId)
                        {
                            found = true;
                            selecteditem = i;
                            break;
                        }
                    }
                }

                /* this can't happen! (famous last words)
                if (!found)
                {
                    _d.SaveConcept("", Concept, "", "","",false);
                    RefreshConceptList("");
                    cbConcept.Text = Concept;
                }
                 * */

                if (found)
                {
                    cbConcept.SelectedItem = selecteditem;
                } else {
                    MessageBox.Show("Sorry, there has been a problem can't find the concept to clone");
                }


            }
        }
        
    }
}
