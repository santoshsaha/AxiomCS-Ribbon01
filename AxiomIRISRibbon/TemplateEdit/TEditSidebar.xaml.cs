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
using System.Data;

using System.IO;
using System.Xml;
using System.Windows.Markup;
using HTMLConverter;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Windows.Threading;
using Telerik.Windows.Controls;

using Microsoft.Win32;

namespace AxiomIRISRibbon.TemplateEdit
{
    /// <summary>
    /// Interaction logic for TEditSidebar.xaml
    /// </summary>
    public partial class TEditSidebar : UserControl
    {

        private Data D;
        private bool IsTemplate;
        private string Id;
        private string Name;
        private Word.Document Doc;

        //Data
        private DataTable DTTemplate;
        private DataTable DTClause;
        private DataTable DTClauseXML;
        private DataTable DTElement;

        //the Currently selected clause and element
        private string CurrentConceptId;
        private string CurrentClauseId;
        private string CurrentElementId;

        bool RefreshOnSave;
        bool PickedClause;
        bool ClauseLock;
        bool ForceRereshClauseXML;
        bool LoadedAllowNone;


        public TEditSidebar(Word.Document WordDoc)
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            this.tabDebug.Visibility = System.Windows.Visibility.Hidden;
            if (Globals.ThisAddIn.getDebug())
            {
                this.tabDebug.Visibility = System.Windows.Visibility.Visible;
            }

            this.D = Globals.ThisAddIn.getData();
            this.Doc = WordDoc;

            // Initiatlise the tables to hold the XML
            this.DTClauseXML = new DataTable();
            this.DTClauseXML.TableName = "ClauseXML";
            this.DTClauseXML.Columns.Add(new DataColumn("Id", typeof(String)));
            this.DTClauseXML.Columns.Add(new DataColumn("XML", typeof(String)));

            this.CurrentConceptId = "";
            this.CurrentClauseId = "";
            this.CurrentElementId = "";

            this.RefreshConceptList();
            this.GetDropDowns();

            this.ClauseLock = true;
            this.ForceRereshClauseXML = false;


        }

        public void Refresh()
        {
            this.PopulateTree();
        }

        public void RefreshClause(string clauseid, string templateclauseid)
        {
            this.UpdateClauseDetails(clauseid, templateclauseid);
            this.PopulateTree();
            this.SelectClause(clauseid);
        }

        // for the element refresh clause after an element has been inserted
        // from a clause edit
        public void RefreshElements()
        {
            this.DTElement = null;
            this.PopulateTree();            
        }

        // Class to hold the Properties in the Tree View
        public class TreeProperty : RadTreeViewItem
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }
            public string Label { get; set; }

            public TreeProperty(string id, string name, string proptype)
            {
                this.Id = id;
                this.Name = name;
                this.Type = proptype;

                string imgsrc = "";

                if (this.Type == "Template")
                {
                    imgsrc = "pack://application:,,,/AxiomIRISRibbon;component/Resources/template-small.png";
                    this.Label = "Contract Template - " + name;
                }
                if (this.Type == "Clause")
                {
                    imgsrc = "pack://application:,,,/AxiomIRISRibbon;component/Resources/clause-small.png";
                    this.Label = "Clause - " + name;
                }
                if (this.Type == "Concept")
                {
                    imgsrc = "";
                    this.Label = "Concept - " + name;
                }
                if (this.Type == "Element")
                {
                    imgsrc = "pack://application:,,,/AxiomIRISRibbon;component/Resources/element-small.png";
                    this.Label = "Element - " + name;
                }

                this.Header = this.Label;
                if (imgsrc != "") this.DefaultImageSrc = imgsrc;
                this.IsExpanded = true;

            }

        }


        private void PopulateTreeTest()
        {
            this.Tree.Items.Clear();
            TreeProperty root = new TreeProperty("root", "Template", "Template");

            TreeProperty c1 = new TreeProperty("c1", "Concept1", "Concept");
            c1.Items.Add(new TreeProperty("cl1", "Clause1", "Clause"));
            c1.Items.Add(new TreeProperty("cl2", "Clause2", "Clause"));
            c1.Items.Add(new TreeProperty("cl3", "Clause3", "Clause"));
            root.Items.Add(c1);

            TreeProperty c2 = new TreeProperty("c2", "Concept2", "Concept");
            c2.Items.Add(new TreeProperty("cl1", "Clause1", "Clause"));
            c2.Items.Add(new TreeProperty("cl2", "Clause2", "Clause"));
            c2.Items.Add(new TreeProperty("cl3", "Clause3", "Clause"));
            root.Items.Add(c2);

            this.Tree.Items.Add(root);
        }


        // Get the values for the dropdowns - just need to do this once when we populate the sidebar        
        // [could even cache]
        private void GetDropDowns()
        {

            Globals.ThisAddIn.ProcessingUpdate("Get Drop Downs");

            //Template Type
            DataTable dt = Utility.HandleData(this.D.GetPickListValues("RibbonTemplate__c", "Type__c", false)).dt;

            RadComboBoxItem i = new RadComboBoxItem();
            i.Content = "";
            this.cbTemplateType.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new RadComboBoxItem();
                i.Content = r["Value"].ToString();
                this.cbTemplateType.Items.Add(i);
            }

            DataReturn dr = Utility.HandleData(this.D.GetPickListValues("RibbonElement__c", "Type__c", false));
            if (!dr.success) return;
            dt = dr.dt;

            //Element Type
            this.cbElementType.Items.Clear();
            i = new RadComboBoxItem();
            i.Content = "";
            cbElementType.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new RadComboBoxItem();
                i.Content = r["Value"].ToString();
                this.cbElementType.Items.Add(i);
            }
        }

        // Get the values for the concepts
        // this may change when the template is open which is why it is public and not the same as Get DropDowns
        public void RefreshConceptList()
        {
            Globals.ThisAddIn.ProcessingUpdate("Get Concept List");

            DataReturn dr = Utility.HandleData(this.D.GetConcepts());
            if (!dr.success) return;
            DataTable dt = dr.dt;

            //Do both combos
            this.cbClauseConcept.Items.Clear();
            RadComboBoxItem i = new RadComboBoxItem();
            i.Content = "";
            cbClauseConcept.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new RadComboBoxItem();
                i.Tag = r["Id"].ToString();
                i.Content = r["Name"].ToString();
                cbClauseConcept.Items.Add(i);
            }

            /*
            cbTClauseConcept.Items.Clear();
            i = new RadComboBoxItem();
            i.Content = "";
            this.cbTClauseConcept.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new RadComboBoxItem();
                i.Tag = r["Id"].ToString();
                i.Content = r["Name"].ToString();
                this.cbTClauseConcept.Items.Add(i);
            }
             * */

        }

        // Cache Routines - need to cut down hitting the database
        private void GetTemplateDetails()
        {
            if (this.DTTemplate == null)
            {
                Globals.ThisAddIn.ProcessingUpdate("Get Template Details");
                DataReturn dr = Utility.HandleData(this.D.GetTemplate(this.Id));
                if (!dr.success) return;
                this.DTTemplate = dr.dt;
                this.DTTemplate.TableName = "Template";
            }
        }

        private void GetClauseDetails()
        {
            if (this.DTClause == null)
            {
                Globals.ThisAddIn.ProcessingUpdate("Get Clause Details");
                DataReturn dr = Utility.HandleData(this.D.GetClause(this.Id));
                if (!dr.success) return;
                this.DTClause = dr.dt;
                this.DTClause.TableName = "Clause";
            }
        }


        // Find any matching clauses and update the Xml - this is when another doc updates a clause, need to update in the template         
        public void RefreshMatchClause(string Id, string Xml)
        {
            if (this.IsTemplate) {
                DataRow[] clauses = this.GetClauseRow(Id);
                if (clauses.Length == 1)
                {
                    // get the clause details and update the XML
                    string templateclauseid = clauses[0]["Id"].ToString();
                    this.UpdateClauseDetails(Id, templateclauseid,Xml);
                    this.Refresh();
                    this.SelectClause(Id);
                }
            }
        }


        // we're editing the concept propery AllowNone on each clause (could have a concept tab - that might be better!)
        // but for now if it updates then update all the clauses so we don't have to requery
        private void UpdateConceptAllowNone(string ConceptId, bool? AllowNone)
        {
            if (AllowNone == null) AllowNone = false;
            foreach (DataRow r in this.DTClause.Rows)
            {
                if (r["Clause__r_Concept__r_Id"].ToString() == ConceptId)
                {
                    r["Clause__r_Concept__r_AllowNone__c"] = AllowNone.ToString();
                }
            }
        }

        // call when you don't have the XML - this wipes it so it will load from Database
        private void UpdateClauseDetails(string ClauseId, string TemplateClauseId)
        {
            this.UpdateClauseDetails(ClauseId, TemplateClauseId,"");
        }

        private void UpdateClauseDetails(string ClauseId, string TemplateClauseId,string Xml)
        {

            // remove current data
            DataRow deleterow = null;
            foreach (DataRow r in this.DTClause.Rows)
            {
                if (r["Id"].ToString() == TemplateClauseId)
                {
                    deleterow = r;
                }
            }
            if (deleterow != null) deleterow.Delete();

            // get the new data
            DataReturn dr = Utility.HandleData(this.D.GetTemplateClause(TemplateClauseId));
            if (!dr.success) return;
            if (dr.dt.Rows.Count > 0)
            {
                this.DTClause.ImportRow(dr.dt.Rows[0]);
            }
            else
            {
                Globals.ThisAddIn.ProcessingStop("");
                MessageBox.Show("Problem finding Clause");
                return;
            }

            // get the elements for this clause
            dr = Utility.HandleData(this.D.GetElements(ClauseId));
            if (!dr.success) return;

            // delete out all the elements for this clause
            if (this.DTElement != null)
            {
                for (int i = this.DTElement.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow r = this.DTElement.Rows[i];
                    if (r["Clause__r_Id"].ToString() == ClauseId)
                    {
                        r.Delete();
                    }
                }
            }
            
            foreach (DataRow addrow in dr.dt.Rows)
            {
                this.DTElement.ImportRow(addrow);                
            }

            if (Xml == "")
            {
                // remove the XML - the routine will get it from SF if its not there
                for (int i = this.DTClauseXML.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow r = this.DTClauseXML.Rows[i];
                    if (r["Id"].ToString() == ClauseId)
                    {
                        r.Delete();
                    }
                }

            }
            else
            {
                // update the XML
                foreach (DataRow r in this.DTClauseXML.Rows)
                {
                    if (r["Id"].ToString() == ClauseId)
                    {
                        r["XML"] = Xml;                        
                    }
                }
            }

        }

        private void GetAllClauses()
        {
            if (this.DTClause == null)
            {
                Globals.ThisAddIn.ProcessingUpdate("Get All The Clauses");
                DataReturn dr = Utility.HandleData(this.D.GetTemplateClauses(this.Id, ""));
                if (!dr.success) return;
                this.DTClause = dr.dt;
                this.DTClause.TableName = "Clause";
            }
        }

        private void GetElements(string clauseid)
        {
            if (this.DTElement == null)
            {
                Globals.ThisAddIn.ProcessingUpdate("Get the Elements for this clause");
                DataReturn dr = Utility.HandleData(this.D.GetElements(clauseid));
                if (!dr.success) return;
                this.DTElement = dr.dt;
            }
        }

        private void GetAllElements(string clausefilter)
        {
            if (this.DTElement == null)
            {
                Globals.ThisAddIn.ProcessingUpdate("Get All The Elements");
                DataReturn dr = Utility.HandleData(this.D.GetMultipleClauseElements(clausefilter));
                if (!dr.success) return;
                this.DTElement = dr.dt;
                this.DTElement.TableName = "Element";
            }
        }


        private string GetClauseXML(string ClauseId)
        {
            string xml = "";
            // see if we already have it
            bool found = false;
            foreach (DataRow r in this.DTClauseXML.Rows)
            {
                if (r["Id"].ToString() == ClauseId)
                {
                    found = true;
                    xml = r["XML"].ToString();
                }
            }

            if (!found)
            {

                string filename = Utility.SaveTempFile(ClauseId);
                DataReturn dr = Utility.HandleData(this.D.GetClauseFile(ClauseId, filename));
                if (!dr.success) return "";

                filename = dr.strRtn;

                if (filename == "")
                {
                    xml = "Problem finding clause";
                }
                else
                {
                    //This is the bit that causes the flash - have to open and close the file
                    Globals.ThisAddIn.Application.ScreenUpdating = false;
                    Word.Document doc = Globals.ThisAddIn.Application.Documents.Open(filename, Visible: false);
                    xml = doc.WordOpenXML;
                    
                    //Add to the datatable
                    DataRow rw = this.DTClauseXML.NewRow();
                    rw["Id"] = ClauseId;
                    rw["XML"] = xml;
                    this.DTClauseXML.Rows.Add(rw);

                    var docclose = (Microsoft.Office.Interop.Word._Document)doc;
                    docclose.Close();
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
            }
            return xml;
        }


        private string AddClauseXML(string ClauseId, string Xml)
        {
            string xml = "";
            // see if we already have it
            bool found = false;
            foreach (DataRow r in this.DTClauseXML.Rows)
            {
                if (r["Id"].ToString() == ClauseId)
                {
                    found = true;
                    r["XML"] = Xml;
                }
            }

            if (!found)
            {
                DataRow rw = this.DTClauseXML.NewRow();
                rw["Id"] = ClauseId;
                rw["XML"] = Xml;
                this.DTClauseXML.Rows.Add(rw);
            }
            return xml;
        }

        public void PopulateTree()
        {
            

            Globals.ThisAddIn.ProcessingUpdate("Populate Tree");

            try
            {

                this.Tree.Items.Clear();
                if (this.Id == null) this.Id = Globals.ThisAddIn.GetCurrentDocId();
                string name = "";
                string xml = "";

                if (this.Id != "")
                {

                    string doctype = "";
                    if (Globals.ThisAddIn.isTemplate(this.Doc))
                    {
                        this.IsTemplate = true;
                        doctype = "Template";

                        this.GetTemplateDetails();
                        if (this.DTTemplate.Rows.Count > 0)
                        {
                            name = this.DTTemplate.Rows[0]["Name"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("Sorry - can't find the Template");
                            return;
                        }

                        //Show the Template and The Template Clause tabs
                        this.tabItemTemplate.Visibility = System.Windows.Visibility.Visible;
                        this.tabItemTemplate.IsSelected = true;
                        this.tabItemTClause.Visibility = System.Windows.Visibility.Visible;

                        //Hide the Clause tab
                        this.tabItemClause.Visibility = System.Windows.Visibility.Collapsed;

                        //Hide the Element Add and Element Delete buttons
                        this.btnElementAdd.Visibility = System.Windows.Visibility.Hidden;
                        this.btnElementDelete.Visibility = System.Windows.Visibility.Hidden;

                        //Populate the Template Properties
                        if (this.DTTemplate.DefaultView.Count > 0)
                        {
                            Utility.UpdateForm(new Grid[] { formGridTemplate }, ((DataRowView)this.DTTemplate.DefaultView[0]).Row);
                            Utility.ReadOnlyForm(false, new Grid[] { formGridTemplate });
                            this.btnSave.IsEnabled = false;
                            this.btnCancel.IsEnabled = false;
                        }
                    }
                    else
                    {
                        this.IsTemplate = false;
                        doctype = "Clause";
                        this.CurrentClauseId = this.Id;
                        this.GetClauseDetails();

                        if (this.DTClause.Rows.Count > 0)
                        {
                            name = this.DTClause.Rows[0]["Name"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("Sorry - can't find the Clause");
                            return;
                        }

                        this.CurrentConceptId = this.DTClause.Rows[0]["Concept__r_Id"].ToString();

                        //Hide the Template Tab
                        this.tabItemTemplate.Visibility = System.Windows.Visibility.Collapsed;
                        this.tabItemTClause.Visibility = System.Windows.Visibility.Collapsed;

                        //Show the Clause tab
                        this.tabItemClause.Visibility = System.Windows.Visibility.Visible;
                        this.tabItemClause.IsSelected = true;

                        //Show the Element Add and Element Delete buttons
                        this.btnElementAdd.Visibility = System.Windows.Visibility.Visible;
                        this.btnElementDelete.Visibility = System.Windows.Visibility.Visible;

                        //Populate the Clause Properties
                        Utility.UpdateForm(new Grid[] { formGridClause }, ((DataRowView)this.DTClause.DefaultView[0]).Row);
                        Utility.ReadOnlyForm(false, new Grid[] { formGridClause });
                        btnSave.IsEnabled = false;
                        btnCancel.IsEnabled = false;

                    }

                    TreeProperty treeContract = new TreeProperty(this.Id, name, doctype);
                    TreeProperty treeItem = null;

                    if (this.IsTemplate)
                    {
                        // Get the concepts from the document so the order is correct

                        Globals.ThisAddIn.ProcessingUpdate("Get Concept Order");
                        string conceptorder = Globals.ThisAddIn.GetConceptOrder(this.Doc);

                        // Get all the clauses
                        xml = "";
                        this.GetAllClauses();

                        // Generate a list of the ClauseIds so we can get all the elemets at once short term solution to cut down on the API calls     
                        // [written after - wonder what the long term solution was!]

                        // TODO - build some clean up in here if the concept doesn't appear in the doc we should delete it from the database

                        List<string> clauseids = new List<string>();
                        foreach (DataRow r in this.DTClause.Rows)
                        {
                            if (!clauseids.Contains(r["Clause__r_Id"].ToString())) clauseids.Add(r["Clause__r_Id"].ToString());
                        }
                        string clausefilter = "";
                        foreach (string c in clauseids)
                        {
                            if (clausefilter == "")
                            {
                                clausefilter = "('" + c + "'";
                            }
                            else
                            {
                                clausefilter += ",'" + c + "'";
                            }
                        }
                        if (clausefilter != "") clausefilter += ")";

                        //Get all the elements
                        this.GetAllElements(clausefilter);

                        //Step through in the right order and display tree
                        if (conceptorder != "")
                        {
                            string[] cotags = conceptorder.Split(',');

                            foreach (string cotag in cotags)
                            {

                                string[] conceptdetails = cotag.Split('|');

                                // format is Concept|ConceptId|ClauseId|LastModified - last 2 may not be there
                                string concept = conceptdetails[1];

                                //Populate all the Concepts and Clauses                            
                                DataView dv = new DataView(this.DTClause);
                                dv.RowFilter = "Clause__r_Concept__r_Id='" + concept + "'";
                                dv.Sort = "Order__c";

                                string conceptname = "";
                                string conceptid = "";
                                if (dv.Count > 0)
                                {
                                    conceptname = dv[0]["Clause__r_Concept__r_Name"].ToString();
                                    conceptid = dv[0]["Clause__r_Concept__r_Id"].ToString();

                                    treeItem = new TreeProperty(conceptid, conceptname, "Concept");
                                }
                                else
                                {
                                    //Concept must have been removed.
                                    Globals.ThisAddIn.RemoveConcept(this.Doc, concept);
                                }



                                foreach (DataRowView r in dv)
                                {
                                    if (Convert.ToString(r["Clause__r_Concept__r_Id"]) != conceptid)
                                    {
                                        //Push the old one  
                                        treeContract.Items.Add(treeItem);

                                        //Create the new one
                                        conceptid = Convert.ToString(r["Clause__r_Concept__r_Id"]);
                                        conceptname = r["Clause__r_Concept__r_Name"].ToString();
                                        treeItem = new TreeProperty(conceptid, conceptname, "Concept");
                                    }


                                    // Get the XML for this clause                                
                                    // ok trying to do this IF the clause in the template matches the clause 
                                    // then just use the XML in the content control so we don't have to go get it

                                    // TODO this does as intended but still has to hit the db if the clause isn't the one selected
                                    // can we stash the xml of the other choices in a hidden bit of the doc?

                                    string clauseid = Convert.ToString(r["Clause__r_Id"]);
                                    string lastmodified = Convert.ToString(r["Clause__r_LastModifiedDate"]);
                                    lastmodified = lastmodified.Substring(0, 16);

                                    xml = "";
                                    bool selectedclause = false;
                                    if (conceptorder != "")
                                    {
                                        cotags = conceptorder.Split(',');
                                        foreach (string cotag1 in cotags)
                                        {
                                            conceptdetails = cotag1.Split('|');
                                            if (conceptdetails.Length > 3)
                                            {
                                                // add a special case for Import - if the tag says 4 zeros then its the uptodate version
                                                if (conceptdetails[2] == clauseid && (lastmodified == conceptdetails[3] || conceptdetails[3] == "0000"))
                                                {
                                                    selectedclause = true;
                                                }
                                            }
                                        }
                                    }

                                    // take the answer from the database if the selection and lastchanged match and we aren't being told to
                                    // (refresh button sets the ForceRefresh flag)
                                    if (selectedclause && (!this.ForceRereshClauseXML))
                                    {
                                        Globals.ThisAddIn.ProcessingUpdate("Take XML from Doc");
                                        // get the xml from the doc
                                        xml = Globals.ThisAddIn.GetTemplateClauseXML(this.Doc, conceptid);
                                        this.AddClauseXML(clauseid, xml);

                                    }

                                    if (xml == "")
                                    {
                                        Globals.ThisAddIn.ProcessingUpdate("Get the Clause Template File from SF for:" + r["Clause__r_Name"].ToString());
                                        xml = this.GetClauseXML(clauseid);
                                    }




                                    // ---------------------


                                    // Globals.ThisAddIn.ProcessingUpdate("Get the Clause Template File from SF for:" + r["Clause__r_Name"].ToString());
                                    // xml = this.GetClauseXML(clauseid);

                                    TreeProperty tpClause = new TreeProperty(Convert.ToString(r["Clause__r_Id"]), r["Clause__r_Name"].ToString(), "Clause");

                                    //Add in the elements for this clause - filter the full list for this clause                                    
                                    DataRow[] elements = this.DTElement.Select("Clause__r_Id='" + Convert.ToString(r["Clause__r_Id"]) + "'");

                                    if (elements.Length > 0)
                                    {
                                        Globals.ThisAddIn.ProcessingUpdate("Get the Elements for:" + r["Clause__r_Name"].ToString());
                                        foreach (DataRow er in elements)
                                        {
                                            tpClause.Items.Add(new TreeProperty(Convert.ToString(er["Element__r_Id"]), er["Element__r_Name"].ToString(), "Element"));
                                        }
                                    }

                                    treeItem.Items.Add(tpClause);
                                }

                                if (dv.Count > 0) treeContract.Items.Add(treeItem);
                            }
                        }
                    }
                    else
                    {
                        this.GetElements(this.Id);
                        //Add in the elements for this clause

                        if (this.DTElement.Rows.Count > 0)
                        {
                            foreach (DataRow er in this.DTElement.Rows)
                            {
                                treeContract.Items.Add(new TreeProperty(Convert.ToString(er["Element__r_Id"]), er["Element__r_Name"].ToString(), "Element"));
                            }
                        }


                    }

                    this.Tree.ExpandAll();
                    this.Tree.Items.Add(treeContract);

                    // if this is a clause then select the clause node
                    if (!this.IsTemplate)
                    {
                        this.SelectClause(this.Id);
                    }



                }
            }
            catch (Exception e)
            {
                string errormessage = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) errormessage += " " + e.InnerException.Message;

                MessageBox.Show(errormessage);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }

        }


        private DataRow[] GetTemplateRow(string id)
        {
            return this.DTTemplate.Select("Id='" + id + "'");
        }


        private DataRow[] GetClauseRow(string id)
        {
            if (this.IsTemplate)
            {
                return this.DTClause.Select("Clause__r_Id='" + id + "'");
            }
            else
            {
                return this.DTClause.Select("Id='" + id + "'");
            }
        }

        private string GetClauseXMLFromId(string id)
        {
            string xml = "";
            DataRow[] dr = this.DTClauseXML.Select("Id='" + id + "'");
            if (dr.Length > 0) xml = Convert.ToString(dr[0]["XML"]);
            return xml;
        }


        private void UpdateClauseXML(string id, string XML)
        {
            DataRow[] dr = this.DTClauseXML.Select("Id='" + id + "'");
            if (dr.Length > 0)
            {
                dr[0]["XML"] = XML;
            }
            return;
        }


        private DataRow[] GetElementRow(string id)
        {
            return this.DTElement.Select("Element__r_Id='" + id + "'");
        }

        private void Tree_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            e.Handled = true;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Send,
                new Action(
                    delegate
                    {

                        if ((TreeProperty)((RadTreeView)sender).SelectedItem != null)
                        {
                            TreeProperty selected = ((TreeProperty)((RadTreeView)sender).SelectedItem);

                            string proptype = selected.Type;
                            string name = selected.Name;
                            string id = selected.Id;

                            // Lock the current clause if it isn't already
                            // will loose changes - should really prompt
                            if (!this.ClauseLock)
                            {
                                imgLock.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                                this.ClauseLock = true;
                                Globals.ThisAddIn.UnlockLockTemplateConcept(this.Doc, this.CurrentConceptId, true);
                            }

                            if (proptype == "Template")
                            {
                                this.tabItemTemplate.IsSelected = true;
                            }
                            else if (proptype == "Concept")
                            {
                                if (!this.PickedClause) Globals.ThisAddIn.SelectContractTemplatesConcept(this.Doc,id);
                                this.PickedClause = false;

                            }
                            else if (this.IsTemplate && proptype == "Clause")
                            {
                                this.PickedClause = true;
                                this.tabItemTClause.IsSelected = true;

                                //Select the concept
                                DataRow clause = GetClauseRow(selected.Id)[0];

                                this.CurrentClauseId = selected.Id;

                                string conceptid = Convert.ToString(clause["Clause__r_Concept__r_Id"]);
                                Globals.ThisAddIn.SelectContractTemplatesConcept(this.Doc,conceptid);

                                if (Globals.ThisAddIn.getDebug())
                                {
                                    this.tbDebugClauseId.Text = this.CurrentClauseId;
                                    this.tbDebugConceptId.Text = conceptid;
                                }

                                this.CurrentConceptId = conceptid;
                                this.LoadedAllowNone = Convert.ToBoolean(clause["Clause__r_Concept__r_AllowNone__c"]);

                                //Update any clauses in the doc with the select clause
                                string xml = this.GetClauseXMLFromId(selected.Id);
                                string lastmodified = Convert.ToString(clause["Clause__r_LastModifiedDate"]);
                                lastmodified = lastmodified.Substring(0, 16);

                                Globals.ThisAddIn.UpdateContractTemplatesConcept(this.Doc, conceptid, id, xml, lastmodified);
                                ((TreeProperty)((RadTreeView)sender).SelectedItem).IsSelected = true;

                                // Populate the Clause Properties
                                Utility.UpdateForm(new Grid[] { this.formGridTClause }, clause);
                                Utility.ReadOnlyForm(false, new Grid[] { this.formGridTClause });
                                this.btnSave.IsEnabled = false;
                                this.btnCancel.IsEnabled = false;

                                // Update the Playbook Links to indicate if there is data and update the hover
                                string ClientPlayBook = Convert.ToString(clause["Clause__r_Concept__r_PlayBookClient__c"]);
                                string InfoPlayBook = Convert.ToString(clause["Clause__r_Concept__r_PlayBookInfo__c"]);

                                this.UpdatePlaybookButton("Client",ClientPlayBook);
                                this.UpdatePlaybookButton("Info", InfoPlayBook);

                            }
                            else if (!this.IsTemplate && proptype == "Clause")
                            {
                                ((TreeProperty)((RadTreeView)sender).SelectedItem).IsSelected = true;
                                tabItemClause.IsSelected = true;

                                DataRow clause = GetClauseRow(selected.Id)[0];

                                //Populate the Clause Properties
                                Utility.UpdateForm(new Grid[] { this.formGridClause }, clause);
                                Utility.ReadOnlyForm(false, new Grid[] { this.formGridClause });
                                this.btnSave.IsEnabled = false;
                                this.btnCancel.IsEnabled = false;

                                // Update the Playbook Links to indicate if there is data and update the hover
                                string ClientPlayBook = Convert.ToString(clause["Concept__r_PlayBookClient__c"]);
                                string InfoPlayBook = Convert.ToString(clause["Concept__r_PlayBookInfo__c"]);

                                this.UpdatePlaybookButton("Client", ClientPlayBook);
                                this.UpdatePlaybookButton("Info", InfoPlayBook);

                            }
                            else if (proptype == "Element")
                            {
                                tabItemElement.IsSelected = true;
                                this.CurrentElementId = selected.Id;
                                DataRow element = GetElementRow(selected.Id)[0];

                                if (this.IsTemplate)
                                {
                                    //Select the clause
                                    //if parent is not the current clause then select it as the current clause                                
                                    TreeProperty parent = (TreeProperty)selected.Parent;
                                    if (this.CurrentClauseId != "" && parent.Id != this.CurrentClauseId)
                                    {
                                        //Select the clause and update values
                                        this.CurrentClauseId = parent.Id;
                                        //Populate the Clause Properties
                                        DataRow clause = GetClauseRow(this.CurrentClauseId)[0];
                                        Utility.UpdateForm(new Grid[] { formGridTClause }, clause);
                                        btnSave.IsEnabled = false;
                                        btnCancel.IsEnabled = false;

                                        string conceptid = Convert.ToString(clause["Clause__r_Concept__r_Id"]);
                                        Globals.ThisAddIn.SelectContractTemplatesConcept(this.Doc,conceptid);
                                        string xml = GetClauseXML(parent.Id);                                        
                                        Globals.ThisAddIn.UpdateContractTemplatesConcept(this.Doc, conceptid, parent.Id, xml,"");
                                    }
                                }

                                //Select the element
                                //Globals.ThisAddIn.SelectElements(id);
                                ((TreeProperty)((RadTreeView)sender).SelectedItem).IsSelected = true;

                                //Populate the Element Properties
                                Utility.UpdateForm(new Grid[] { this.formGridElement }, element);
                                Utility.ReadOnlyForm(false, new Grid[] { this.formGridElement });
                                this.btnSave.IsEnabled = false;
                                this.btnCancel.IsEnabled = false;

                            }
                        }

                    }));
        }



        public void TreeSelect(TreeProperty tree, string id, string proptype)
        {
            // Check root
            if (tree.Type == proptype && (tree.Id == id))
            {
                tree.IsSelected = true;
                return;
            }

            //Step through the tree and select the matching node
            foreach (TreeProperty item in tree.Items)
            {
                if (item.Type == proptype && (item.Id == id))
                {
                    item.IsSelected = true;
                    return;
                }
                else
                {
                    if (item.Items.Count > 0)
                    {
                        this.TreeSelect((TreeProperty)item, id, proptype);
                    }
                }
            }
        }

        public void SelectConcept(string conceptid)
        {            
            if (this.Tree.Items.Count > 0)
            {
                TreeSelect((TreeProperty)this.Tree.Items[0], conceptid, "Concept");
            }

        }

        public void SelectClause(string clauseid)
        {
            if (this.Tree.Items.Count > 0)
            {
                TreeSelect((TreeProperty)this.Tree.Items[0], clauseid, "Clause");
            }
        }


        public void SelectElement(string id)
        {
            if (this.Tree.Items.Count > 0)
            {
                TreeSelect((TreeProperty)this.Tree.Items[0], id, "Element");
            }
        }

        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.ProcessingStart("Reload Tree");
            this.Refresh();
            Globals.ThisAddIn.ProcessingStop("Finished");
        }



        // Template/Clause/Element form events

        private void FormTextChanged(object sender, TextChangedEventArgs e)
        {
            this.btnSave.IsEnabled = true;
            this.btnCancel.IsEnabled = true;
        }

        private void FormSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.btnSave.IsEnabled = true;
            this.btnCancel.IsEnabled = true;

            RadComboBox r = (RadComboBox)sender;
            if (r.Name == "cbClauseConcept")
            {
                this.RefreshOnSave = true;
            }

        }

        private void FormCheckBoxClick(object sender, RoutedEventArgs e)
        {
            this.btnSave.IsEnabled = true;
            this.btnCancel.IsEnabled = true;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            this.FormSave();
        }

        public void FormSave()
        {            
            RadTabItem selectedtab = (RadTabItem)this.tcEdit.SelectedItem;
            if (btnSave.IsEnabled)
            {
                if (this.tabItemTemplate == selectedtab)
                {
                    if (this.tbTemplateName.Text == "")
                    {
                        MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                        return;
                    }

                    //Update from the form
                    DataRow template = this.GetTemplateRow(this.Id)[0];
                    Utility.UpdateRow(new Grid[] { this.formGridTemplate }, template);

                    //Save the values the document
                    Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                    DataReturn dr = Utility.HandleData(this.D.SaveTemplate(template));
                    if (!dr.success) return;
                    this.Id = dr.id;

                    this.SaveDoc();

                    this.btnSave.IsEnabled = false;
                    this.btnCancel.IsEnabled = false;
                }
                else if (this.tabItemTClause == selectedtab)
                {
                    //check the required fields
                    if (this.tbTClauseName.Text == "")
                    {
                        MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                        return;
                    }

                    //Update from the form - problem is this has been loaded from the
                    //templateclause table not the clause table
                    DataRow clause = this.GetClauseRow(this.CurrentClauseId)[0];
                    Utility.UpdateRow(new Grid[] { this.formGridTClause }, clause);

                    DataReturn dr = Utility.HandleData(this.D.SaveClauseFromTemplateClause(clause));
                    if (!dr.success) return;
                    string clauseid = dr.id;

                    // Save the order and the default selection
                    dr = Utility.HandleData(this.D.UpdateTemplateClause(clause["Id"].ToString(), tbClauseOrder.Text,tbClauseDefault.Text));

                    // If unlocked save the XML in the content control as the attachment
                    if (!this.ClauseLock)
                    {
                        Globals.ThisAddIn.ProcessingUpdate("Unlocked, Update Clause File");

                        string text = Globals.ThisAddIn.GetTemplateClauseText(this.Doc, this.CurrentConceptId);
                        string xml = Globals.ThisAddIn.GetTemplateClauseXML(this.Doc, this.CurrentConceptId);

                        // Update the XML in the cache
                        this.UpdateClauseXML(clauseid, xml);

                        // Save the Attachment to Salesforce
                        string clausefilename = Utility.SaveTempFile(clauseid);

                        Word.Document scratch = Globals.ThisAddIn.Application.Documents.Add(Visible: false);
                        scratch.Content.InsertXML(xml);
                        scratch.SaveAs2(FileName: clausefilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                        var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                        docclosescratch.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                        // Now save the file
                        Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                        dr = Utility.HandleData(this.D.SaveClauseFile(clauseid, text, clausefilename));
                        if (!dr.success) return;

                        // this returns the lastmodified date - need to update the concept tag
                        string lastmodifieddate = dr.strRtn;                        
                        lastmodifieddate = lastmodifieddate.Substring(0, 16);
                        Globals.ThisAddIn.UpdateContractTemplatesConceptTag(this.Doc, this.CurrentConceptId, clauseid, lastmodifieddate);

                        // and update the cache
                        clause = GetClauseRow(this.CurrentClauseId)[0];
                        clause["Clause__r_LastModifiedDate"] = lastmodifieddate;

                        // and lock
                        imgLock.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                        this.ClauseLock = true;
                        Globals.ThisAddIn.UnlockLockTemplateConcept(this.Doc, this.CurrentConceptId, true);

                        this.btnSave.IsEnabled = true;
                        this.btnCancel.IsEnabled = true;

                        // need to save the template to make sure it has the clause change
                        if (!this.Doc.Saved)
                        {
                            this.SaveDoc();
                        }
                    }

                    // if the allow none on the concept has changed then save that
                    if (this.cbAllowNone.IsChecked != this.LoadedAllowNone)
                    {
                        dr = Utility.HandleData(this.D.UpdateConceptAllowNone(clause["Clause__r_Concept__r_Id"].ToString(), this.cbAllowNone.IsChecked));
                        if (!dr.success) return;

                        // also update the other clauses with this concept
                        this.UpdateConceptAllowNone(clause["Clause__r_Concept__r_Id"].ToString(), this.cbAllowNone.IsChecked);
                    }




                    if (this.RefreshOnSave)
                    {
                        Refresh();
                    }


                    this.btnSave.IsEnabled = false;
                    this.btnCancel.IsEnabled = false;
                    this.RefreshOnSave = false;
                }
                else if (this.tabItemClause == selectedtab)
                {
                    //check the required fields
                    if (this.tbClauseName.Text == "")
                    {
                        MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                        return;
                    }

                    //Update from the form
                    DataRow clause = GetClauseRow(this.CurrentClauseId)[0];
                    Utility.UpdateRow(new Grid[] { this.formGridClause }, clause);
                    DataReturn dr = Utility.HandleData(this.D.SaveClause(clause));
                    if (!dr.success) return;
                    string clauseid = dr.id;

                    this.btnSave.IsEnabled = false;
                    this.btnCancel.IsEnabled = false;
                }
                else if (this.tabItemElement == selectedtab)
                {
                    //check the required fields
                    if (this.tbElementName.Text == "")
                    {
                        MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                        return;
                    }
                    if (this.cbElementType.Text == "")
                    {
                        MessageBox.Show("Type field is required", "Problem", MessageBoxButton.OK);
                        return;
                    }

                    //Update from the form
                    DataRow element = GetElementRow(this.CurrentElementId)[0];
                    Utility.UpdateRow(new Grid[] { this.formGridElement }, element);


                    DataReturn dr = Utility.HandleData(this.D.SaveElementFromClauseElement(element));
                    if (!dr.success) return;
                    string elementid = dr.id;

                    //Save the order
                    int intord;
                    if (Int32.TryParse(this.tbElementOrder.Text, out intord))
                    {
                        dr = Utility.HandleData(this.D.UpdateClauseElementOrder(element["Id"].ToString(), this.tbElementOrder.Text));
                        if (!dr.success) return;
                    }


                    this.btnSave.IsEnabled = false;
                    this.btnCancel.IsEnabled = false;
                }
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {

            RadTabItem selectedtab = (RadTabItem)this.tcEdit.SelectedItem;

            if (this.tabItemTemplate==selectedtab)
            {
                //Undo changes
                DataRow template = this.GetTemplateRow(this.Id)[0];
                Utility.UpdateForm(new Grid[] { this.formGridTemplate }, template);
                this.btnSave.IsEnabled = false;
                this.btnCancel.IsEnabled = false;
            }
            else if (this.tabItemTClause == selectedtab)
            {
                DataRow clause = GetClauseRow(this.CurrentClauseId)[0];
                Utility.UpdateForm(new Grid[] { this.formGridTClause }, clause);
                this.btnSave.IsEnabled = false;
                this.btnCancel.IsEnabled = false;
            }
            else if (this.tabItemClause == selectedtab)
            {
                DataRow clause = this.GetClauseRow(this.CurrentClauseId)[0];
                Utility.UpdateForm(new Grid[] { this.formGridClause }, clause);
                this.btnSave.IsEnabled = false;
                this.btnCancel.IsEnabled = false;
            }
            else if (this.tabItemElement == selectedtab)
            {
                DataRow element = GetElementRow(this.CurrentElementId)[0];
                Utility.UpdateForm(new Grid[] { this.formGridElement }, element);
                this.btnSave.IsEnabled = false;
                this.btnCancel.IsEnabled = false;
            }
        }

        private void btnTClauseAdd_Click(object sender, RoutedEventArgs e)
        {
                //Add in Clause
                Clause ucClause = Globals.ThisAddIn.OpenClause(true,true);
                DataRow template = GetTemplateRow(this.Id)[0];
                string templateid = template["Id"].ToString();
                string templatename = template["Name"].ToString();
                ucClause.SetFromTemplate(templateid, templatename, "add");
                ucClause.Show();
            
        }

        private void btnTClauseNew_Click(object sender, RoutedEventArgs e)
        {

            //Add in Clause
            Clause ucClause = Globals.ThisAddIn.OpenClause(true,false);
            DataRow template = GetTemplateRow(this.Id)[0];
            string templateid = template["Id"].ToString();
            string templatename = template["Name"].ToString();
            ucClause.SetFromTemplate(templateid, templatename, "inplace");

            //This is if we are making a NEW Clause -----------------
            Word.Selection sel;
            Word.Document doc;

            doc = Globals.ThisAddIn.Application.ActiveDocument;
            sel = Globals.ThisAddIn.Application.Selection;

            // check if the character before the start of the range is not a new line
            // if not select to the start of the line
            if (sel.Start > 1)
            {
                if (doc.Range(sel.Start - 1, sel.Start).Characters[1].Text != "\r")
                {
                    sel.MoveStart(Word.WdUnits.wdLine, -1);
                }
            }

            // check if the last character in the range is a new line and if not select
            // to the end
            if (sel.Range.Characters.Last.Text != "\r")
            {
                sel.MoveEnd(Word.WdUnits.wdLine, 1);
            }

            //Try and work out the name and the description
            string name = "";
            if (sel.Range.Text != null && sel.Range.Text.Length > 0)
            {
                foreach (Word.Range r in sel.Range.Words)
                {
                    //if its a little word then skip it!
                    if (r == null || r.Text.Length > 3)
                    {
                        name = r.Text;
                        break;
                    }
                }

                string desc = "";
                foreach (Word.Range r in sel.Range.Sentences)
                {
                    //if its a little senetence then skip it!
                    if (r == null || r.Text.Length > 5)
                    {
                        desc = r.Text;
                        if (desc.Length > 20) break;
                    }
                }

                //if the current selection is in a doc that is in a template then this is edit in place
                //pass the doc id so when we create we know what to do

                if (Globals.ThisAddIn.isTemplate(doc)) templateid = Globals.ThisAddIn.GetDocId(doc);
                ucClause.NewClause(name, desc, "", sel.Text, sel.WordOpenXML);
            }
            else
            {


            }
            ucClause.Show();

        }

        private void btnElementAdd_Click(object sender, RoutedEventArgs e)
        {
            //Add in Clause
            Element ucElement = Globals.ThisAddIn.OpenElement();

            Word.Selection sel;
            Word.Document doc;

            doc = Globals.ThisAddIn.Application.ActiveDocument;
            sel = Globals.ThisAddIn.Application.Selection;

            //Try and work out the name and the description
            string name = "";
            if (sel.Range.Text != null && sel.Range.Text.Length > 0)
            {
                name = sel.Range.Text;
                ucElement.NewElement(name, "", sel.Text, sel.WordOpenXML, this.Id, name);
            }
        }

        private void btnTClauseEdit_Click(object sender, RoutedEventArgs e)
        {
            //Get the clause Id
            if (this.CurrentClauseId != null)
            {
                Clause ucClause = Globals.ThisAddIn.OpenClause(false,false);
                ucClause.OpenClause(this.CurrentClauseId);
                ucClause.Hide();
            }
        }

        private void btnTClauseDelete_Click(object sender, RoutedEventArgs e)
        {
            //Remove the link between this template and the clause
            if (this.CurrentClauseId != "")
            {
                DataRow templateclause = this.GetClauseRow(this.CurrentClauseId)[0];


                string templateid = this.Id;
                string templateclauseid = templateclause["Id"].ToString();
                //Are you sure

                MessageBoxResult res = MessageBox.Show("Are you sure?", "Warning", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    return;
                }

                DataReturn dr = Utility.HandleData(this.D.DeleteTemplateClause(templateclauseid));
                if (!dr.success) return;

                // delete from cache
                DataRow deleterow = null;                
                foreach (DataRow r in this.DTClause.Rows)
                {
                    if (r["Id"].ToString() == templateclauseid)
                    {
                        deleterow = r;
                    } 
                }
                if (deleterow != null) deleterow.Delete();
               
                // don't need to remove the concept control if there are none left - the PopulateTree does this
                this.Refresh();
                this.SaveDoc();
                
            }
        }



        private void SaveDoc()
        {
            // the save throws an error intentionally to stop the regular save, just trap
            try
            {
                this.Doc.Save();
            }
            catch (Exception)
            {
            }

        }

        private void btnElementEdit_Click(object sender, RoutedEventArgs e)
        {
            if (this.CurrentElementId != "")
            {

                string causeid = this.Id;
                DataRow element = GetElementRow(this.CurrentElementId)[0];
                string caluseelementid = element["Id"].ToString();
                //Open Element
                Element ucElement = Globals.ThisAddIn.OpenElement();
                ucElement.Open(caluseelementid);

            }
        }

        private void btnElementDelete_Click(object sender, RoutedEventArgs e)
        {
            //delete from the database and the document
            //Remove the link between this caluse and the element

            if (!this.IsTemplate && this.CurrentElementId != "")
            {

                DataRow element = this.GetElementRow(this.CurrentElementId)[0];
                string causeid = this.Id;
                string elementid = element["Element__r_Id"].ToString();
                string clauseelementid = element["Id"].ToString();
                //Are you sure

                MessageBoxResult res = MessageBox.Show("Are you sure?", "Warning", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    return;
                }

                // Delete from Salesforce
                DataReturn dr = Utility.HandleData(this.D.DeleteClauseElement(clauseelementid));
                if (!dr.success) return;

                // Delete from cache
                DataRow deleterow = null;
                foreach (DataRow r in this.DTElement.Rows)
                {
                    if (r["Id"].ToString() == clauseelementid)
                    {
                        deleterow = r;
                    }
                }
                if (deleterow != null) deleterow.Delete();
                this.Refresh();
                
                Globals.ThisAddIn.RemoveElements(this.Doc, elementid);

                SaveDoc();
            }
        }

        private void btnTClauseLock_Click(object sender, RoutedEventArgs e)
        {
            if(CurrentClauseId!=""){

                DataRow clause = GetClauseRow(CurrentClauseId)[0];
                string conceptid = Convert.ToString(clause["Clause__r_Concept__r_Id"]);

            if (this.ClauseLock)
            {                
                imgLock.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/unlocksmall.png", UriKind.Relative));
                this.ClauseLock = false;

                // Unlock the content control
                Globals.ThisAddIn.UnlockLockTemplateConcept(this.Doc,conceptid,false);
                this.btnSave.IsEnabled = true;
                this.btnCancel.IsEnabled = true;

            }
            else
            {             
                imgLock.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                this.ClauseLock = true;

                // lock the content control
                Globals.ThisAddIn.UnlockLockTemplateConcept(this.Doc, conceptid, true);
            }
            }
        }

        private void btnTClauseClone_Click(object sender, RoutedEventArgs e)
        {

            //Get the clause Id
            if (this.CurrentClauseId != null && this.CurrentClauseId != "")
            {
                Clause ucClause = Globals.ThisAddIn.OpenClause(true,false);
                DataRow template = GetTemplateRow(this.Id)[0];
                string templateid = template["Id"].ToString();
                string templatename = template["Name"].ToString();
                ucClause.SetFromTemplate(templateid, templatename, "clone");                

                DataRow clause = GetClauseRow(this.CurrentClauseId)[0];
                
                string name = this.tbTClauseName.Text;
                string desc = this.tbTClauseDescription.Text;

                ucClause.CloneClause(name, desc, this.CurrentConceptId, this.CurrentClauseId);
                ucClause.Show();
            }
        }

        private void btnCreateContract_Click(object sender, RoutedEventArgs e)
        {
            //Create an instance from this template
            if (this.IsTemplate)
            {
                Contract axC = new Contract();
                DataRow template = this.GetTemplateRow(this.Id)[0];
                axC.Open("", this.Id, template["Name"].ToString(), template["PlaybookLink__c"].ToString());
            }
        }

        private void btnPlaybookClient_Click(object sender, RoutedEventArgs e)
        {
            if(this.CurrentClauseId!=""){
                DataRow clause = this.GetClauseRow(this.CurrentClauseId)[0];
                string html = clause["Clause__r_Concept__r_PlayBookClient__c"].ToString();
                Playbook p = new Playbook();
                p.Open(this,this.CurrentConceptId, html, "Client");
                p.Show();
            }
        }

        private void btnPlaybookInfo_Click(object sender, RoutedEventArgs e)
        {
            if (this.CurrentClauseId != "")
            {
                DataRow clause = this.GetClauseRow(this.CurrentClauseId)[0];
                string html = clause["Clause__r_Concept__r_PlayBookInfo__c"].ToString();
                Playbook p = new Playbook();
                p.Open(this,this.CurrentConceptId, html, "Info");
                p.Show();
            }
        }

        private void btnClausePlaybookClient_Click(object sender, RoutedEventArgs e)
        {
            if (this.CurrentClauseId != "")
            {
                DataRow clause = this.GetClauseRow(this.CurrentClauseId)[0];
                string html = clause["Concept__r_PlayBookClient__c"].ToString();
                Playbook p = new Playbook();
                p.Open(this, this.CurrentConceptId, html, "Client");
                p.Show();
            }
        }

        private void btnClausePlaybookInfo_Click(object sender, RoutedEventArgs e)
        {
            if (this.CurrentClauseId != "")
            {
                DataRow clause = this.GetClauseRow(this.CurrentClauseId)[0];
                string html = clause["Concept__r_PlayBookInfo__c"].ToString();
                Playbook p = new Playbook();
                p.Open(this, this.CurrentConceptId, html, "Info");
                p.Show();
            }
        }


        // Update each of the clause lines with the update concept play book info
        // used after an edit to update the values cached in the sidebar - playbook edit
        // looks after updating Salesforce
        public void UpdateCachePlaybook(string ConceptId, string PBType, string html)
        {
            if (this.IsTemplate)
            {
                foreach (DataRow r in this.DTClause.Rows)
                {
                    if (r["Clause__r_Concept__r_Id"].ToString() == ConceptId)
                    {
                        if (PBType == "Client")
                        {
                            r["Clause__r_Concept__r_PlayBookClient__c"] = html;
                        }
                        else if (PBType == "Info")
                        {
                            r["Clause__r_Concept__r_PlayBookInfo__c"] = html;
                        }
                    }
                }
            }
            else
            {
                foreach (DataRow r in this.DTClause.Rows)
                {
                    if (r["Concept__r_Id"].ToString() == ConceptId)
                    {
                        if (PBType == "Client")
                        {
                            r["Concept__r_PlayBookClient__c"] = html;
                        }
                        else if (PBType == "Info")
                        {
                            r["Concept__r_PlayBookInfo__c"] = html;
                        }
                    }
                }
            }

            // if this is the one that is currenly selected then update the button
            if (this.CurrentConceptId == ConceptId)
            {
                this.UpdatePlaybookButton(PBType, html);
            }
        }

        // Get Footnotes
        public string GetFootnotes()
        {
            return Globals.ThisAddIn.GetFootnotes(this.Doc,this.CurrentConceptId);
        }



        public void UpdatePlaybookButton(string PBType, string html)
        {
            if (IsTemplate)
            {
                if (PBType == "Info")
                {
                    if (html.Trim() == "")
                    {
                        this.btnPlaybookInfo.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                        this.btnPlaybookInfo.ToolTip = null;
                    }
                    else
                    {
                        this.btnPlaybookInfo.Foreground = (SolidColorBrush)new BrushConverter().ConvertFromString("Blue");
                        this.btnPlaybookInfo.ToolTip = ConvertHTMLToToolTip(html);
                    }
                }

                else if (PBType == "Client")
                {
                    if (html.Trim() == "")
                    {
                        this.btnPlaybookClient.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                        this.btnPlaybookClient.ToolTip = null;
                    }
                    else
                    {
                        this.btnPlaybookClient.Foreground = (SolidColorBrush)new BrushConverter().ConvertFromString("Blue");
                        this.btnPlaybookClient.ToolTip = ConvertHTMLToToolTip(html);
                    }
                }
            }
            else
            {
                if (PBType == "Info")
                {
                    if (html.Trim() == "")
                    {
                        this.btnClausePlaybookInfo.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                        this.btnClausePlaybookInfo.ToolTip = null;
                    }
                    else
                    {
                        this.btnClausePlaybookInfo.Foreground = (SolidColorBrush)new BrushConverter().ConvertFromString("Blue");
                        this.btnClausePlaybookInfo.ToolTip = ConvertHTMLToToolTip(html);
                    }
                }

                else if (PBType == "Client")
                {
                    if (html.Trim() == "")
                    {
                        this.btnClausePlaybookClient.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                        this.btnClausePlaybookClient.ToolTip = null;
                    }
                    else
                    {
                        this.btnClausePlaybookClient.Foreground = (SolidColorBrush)new BrushConverter().ConvertFromString("Blue");
                        this.btnClausePlaybookClient.ToolTip = ConvertHTMLToToolTip(html);
                    }
                }
            }
        }


        private string ConvertHTMLToString(string html)
        {
            string str = "";
            if (html.Trim() != "")
            {
                // -- Get the HTML as text - convert through xaml - must be a better way!                        
                StringReader stringReader = new StringReader(HtmlToXamlConverter.ConvertHtmlToXaml(html, true));
                XmlReader xmlReader = XmlReader.Create(stringReader);
                FlowDocument fdoc = (FlowDocument)XamlReader.Load(xmlReader);

                TextRange tr = new TextRange(fdoc.ContentStart, fdoc.ContentEnd);

                MemoryStream ms = new MemoryStream();
                ms = new MemoryStream();
                tr.Save(ms, DataFormats.Text);
                ms.Seek(0, SeekOrigin.Begin);
                string s = new StreamReader(ms).ReadToEnd();

                str = s;
            }
            return str;
        }

        private ToolTip ConvertHTMLToToolTip(string html)
        {
            ToolTip tt = null;
            if (html.Trim() != "")
            {
                // -- Get the HTML as text - convert through xaml - must be a better way!                        
                StringReader stringReader = new StringReader(HtmlToXamlConverter.ConvertHtmlToXaml(html, true));
                XmlReader xmlReader = XmlReader.Create(stringReader);
                FlowDocument fd = (FlowDocument)XamlReader.Load(xmlReader);

                RichTextBox rt = new RichTextBox();
                rt.Document = fd;
                rt.Background = Brushes.Transparent;
                rt.BorderThickness = new Thickness(0);

                tt = new ToolTip();
                tt.MaxWidth = 300;
                tt.MaxHeight = 200;
                tt.Content = rt;
                tt.BorderThickness = new Thickness(1);

            }
            return tt;
        }

        private void btnTemplatePlaybookLink_Click(object sender, RoutedEventArgs e)
        {
            // open a browser with the link
            if (this.tbTemplatePlaybook.Text != "")
            {
                System.Diagnostics.Process.Start(this.tbTemplatePlaybook.Text);
            }
        }

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
                       
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;

            // check if the character before the start of the range is not a new line
            // if not select to the start of the line
            if (sel.Start > 1)
            {
                if (Doc.Range(sel.Start - 1, sel.Start).Characters[1].Text != "\r")
                {
                    sel.MoveStart(Word.WdUnits.wdLine, -1);
                }
            }

            // check if the last character in the range is a new line and if not select
            // to the end

            if (sel.Range.Characters.Last.Text != "\r")
            {
                sel.MoveEnd(Word.WdUnits.wdLine, 1);
            }
            
        }

        private void btnTClauseReload_Click(object sender, RoutedEventArgs e)
        {
            // Refresh with the XML set to blank to get the tree to reload the clause from Sforce
            if (this.CurrentClauseId != null)
            {
                this.ForceRereshClauseXML = true;
                Globals.ThisAddIn.ProcessingStart("Get Clause from Salesforce");
                
                if (this.IsTemplate)
                {
                    DataRow[] clauses = this.GetClauseRow(CurrentClauseId);
                    if (clauses.Length == 1)
                    {
                        // get the clause details and update the XML
                        string templateclauseid = clauses[0]["Id"].ToString();
                        this.UpdateClauseDetails(CurrentClauseId, templateclauseid, "");
                        this.Refresh();
                        this.SelectClause(CurrentClauseId);
                    }
                }


                Globals.ThisAddIn.ProcessingStop("");
                this.ForceRereshClauseXML = false;
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {

            // so whats the plan! 
            // for now don't make it human readable i.e. not a marked up doc - going to be too clumsy, might come back to
            // so leave in the controls and write the metadata and the non-shown clauses to custom xml parts

            // leave in the ids as they are - import will sort them out
            try
            {
                Globals.ThisAddIn.ProcessingStart("Import Template");

                if (this.IsTemplate)
                {

                    Globals.ThisAddIn.ProcessingUpdate("Copy Template");
                    Word.Document template = Globals.ThisAddIn.Application.ActiveDocument;
                    Word.Document export = Globals.ThisAddIn.Application.Documents.Add();

                    Word.Range source = template.Range(template.Content.Start, template.Content.End);
                    export.Range(export.Content.Start).InsertXML(source.WordOpenXML);

                    Globals.ThisAddIn.ProcessingUpdate("Update Template Id");
                    Globals.ThisAddIn.AddDocId(export, "ExportTemplate", this.Id);

                    // now get the meta data and store it as custom xml parts
                    Globals.ThisAddIn.ProcessingUpdate("Add Data to Dataset");
                    DataSet ds = new DataSet();
                    ds.Namespace = "http://www.axiomlaw.com/irisribbon";
                    Globals.ThisAddIn.ProcessingUpdate("Template");
                    ds.Tables.Add(this.DTTemplate);
                    Globals.ThisAddIn.ProcessingUpdate("Clause");
                    ds.Tables.Add(this.DTClause);
                    Globals.ThisAddIn.ProcessingUpdate("ClauseXML");
                    ds.Tables.Add(this.DTClauseXML);
                    Globals.ThisAddIn.ProcessingUpdate("Elements");
                    ds.Tables.Add(this.DTElement);

                    Globals.ThisAddIn.ProcessingUpdate("Serialise Data to XML");
                    string xmldata = "";
                    using (StringWriter stringWriter = new StringWriter())
                    {
                        ds.WriteXml(new XmlTextWriter(stringWriter));
                        xmldata = stringWriter.ToString();
                    };

                    xmldata = AxiomIRISRibbon.Utility.CleanUpXML(xmldata);

                    Globals.ThisAddIn.ProcessingUpdate("Save as XML Part");
                    Office.CustomXMLPart data = export.CustomXMLParts.Add(xmldata);

                    export.Activate();

                    Globals.ThisAddIn.ProcessingUpdate("Save As ...");

                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.Filter = "Word Document (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docx";
                    dlg.FileName = "ExportTemplate-" + this.tbTemplateName.Text.Replace(" ", "");
                    Nullable<bool> result = dlg.ShowDialog();
                    if (result == true)
                    {
                        export.SaveAs2(dlg.FileName);
                    }

                    Globals.ThisAddIn.ProcessingStop("Finished");

                }
            }
            catch (Exception ex)
            {
                string message = "Sorry there has been an error - " + ex.Message;
                if (ex.InnerException != null) message += " " + ex.InnerException.Message;
                MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }
            
        }

        public void Import()
        {
            // do the thing!
            try
            {
                Globals.ThisAddIn.ProcessingStart("Import Template");

                Office.CustomXMLPart data = null;
                foreach (Office.CustomXMLPart part in this.Doc.CustomXMLParts)
                {
                    if (part.NamespaceURI == "http://www.axiomlaw.com/irisribbon")
                    {
                        data = part;
                    }
                }

                if (data != null)
                {

                    Globals.ThisAddIn.ProcessingStart("Read Data from Template");

                    DataSet ds = new DataSet();

                    System.IO.StringReader xmlSR = new System.IO.StringReader(data.DocumentElement.XML);

                    ds.ReadXml(xmlSR, XmlReadMode.Auto);

                    this.DTTemplate = ds.Tables["Template"];
                    this.DTClause = ds.Tables["Clause"];
                    this.DTClauseXML = ds.Tables["ClauseXML"];
                    this.DTElement = ds.Tables["Element"];


                    // ok we have the data from the exported template 
                    // now we need to save everything and update the ids

                    // keep list of the old ids
                    string OldTemplateId = this.Id;
                    Dictionary<string, string> ConceptIdMapping = new Dictionary<string, string>();
                    Dictionary<string, string> ClauseIdMapping = new Dictionary<string, string>();
                    Dictionary<string, string> ElementIdMapping = new Dictionary<string, string>();

                    // *TODO* Check for NAME amd ID clashes - also ask if 
                    // they want to create new clauses or update and if the want to create new elements or update


                    // first the template
                    string newtemplateid = "";
                    string newtemplatename = this.DTTemplate.Rows[0]["Name"].ToString();

                    // look for a Template with a matching name
                    DataReturn checktemplate = Utility.HandleData(this.D.CheckTemplate(this.DTTemplate.Rows[0]["Name"].ToString()));
                    if (!checktemplate.success)
                    {
                        return;
                    }

                    // check if we have a match and if so warn them
                    bool updatetemplate = false;
                    if (checktemplate.dt.Rows.Count > 0)
                    {
                        MessageBoxResult res = MessageBox.Show("There is already a template called '" + this.DTTemplate.Rows[0]["Name"].ToString() + "'  Would you like to update from this template?", "Warning", MessageBoxButton.OKCancel);
                        if (res == MessageBoxResult.Cancel)
                        {
                            Globals.ThisAddIn.ProcessingStop("Finished");
                            // hide the sidebar
                            Globals.ThisAddIn.ShowTaskPane(false);
                            return;
                        }
                        newtemplateid = checktemplate.dt.Rows[0][0].ToString();
                        updatetemplate = true;
                    }

                    // Update the global Id and dataset to either the template to update
                    // OR blank if we are creating a new one
                    this.Id = newtemplateid;
                    this.DTTemplate.Rows[0]["Id"] = newtemplateid;

                    if (newtemplateid == "") Globals.ThisAddIn.ProcessingUpdate("Create New Template as " + newtemplatename);

                    DataReturn dr = Utility.HandleData(this.D.SaveTemplate(this.DTTemplate.Rows[0]));
                    if (!dr.success)
                    {
                        MessageBox.Show("There has been an issue - template will not be imported");
                        return;
                    }

                    // remember the new id
                    this.Id = dr.id;
                    this.DTTemplate.Rows[0]["Id"] = dr.id;

                    // set the template id in the doc to get things to work
                    Globals.ThisAddIn.AddDocId(this.Doc, "ContractTemplate", this.Id);

                    // update the form
                    Utility.UpdateForm(new Grid[] { formGridTemplate }, ((DataRowView)this.DTTemplate.DefaultView[0]).Row);
                    Utility.ReadOnlyForm(false, new Grid[] { formGridTemplate });
                    this.btnSave.IsEnabled = false;
                    this.btnCancel.IsEnabled = false;

                    // now step through the clauses and create any that aren't there
                    // step through in the order of the concepts in the template
                    Globals.ThisAddIn.ProcessingUpdate("Get Concept Order");
                    string conceptorder = Globals.ThisAddIn.GetConceptOrder(this.Doc);

                    if (conceptorder != "")
                    {
                        string[] cotags = conceptorder.Split(',');

                        foreach (string cotag in cotags)
                        {

                            string[] conceptdetails = cotag.Split('|');

                            // format is Concept|ConceptId|ClauseId|LastModified - last 2 may not be there
                            string concept = conceptdetails[1];

                            //Populate all the Concepts and Clauses                            
                            DataView dv = new DataView(this.DTClause);
                            dv.RowFilter = "Clause__r_Concept__r_Id='" + concept + "'";
                            dv.Sort = "Order__c";

                            bool first = true;
                            foreach (DataRowView r in dv)
                            {
                                if (first)
                                {
                                    // check if we already have the concept name
                                    DataReturn checkconcept = Utility.HandleData(this.D.CheckConcept(r["Clause__r_Concept__r_Name"].ToString()));
                                    if (!checkconcept.success)
                                    {
                                        return;
                                    }

                                    string ConceptIdOld = r["Clause__r_Concept__r_Id"].ToString();

                                    string ConceptIdNew = "";
                                    if (checkconcept.dt.Rows.Count > 0)
                                    {
                                        ConceptIdNew = checkconcept.dt.Rows[0][0].ToString();
                                    }

                                    string ConceptName = r["Clause__r_Concept__r_Name"].ToString();
                                    string ConceptDescription = r["Clause__r_Concept__r_Description__c"].ToString();
                                    string ConceptPlayBookInfo = r["Clause__r_Concept__r_PlayBookInfo__c"].ToString();
                                    string ConceptPlayBookClient = r["Clause__r_Concept__r_PlayBookClient__c"].ToString();
                                    bool ConceptAllowNone = Convert.ToBoolean(r["Clause__r_Concept__r_AllowNone__c"]);


                                    Globals.ThisAddIn.ProcessingUpdate(ConceptIdNew != "" ? "Updating Concept " + ConceptName : "Save Concept as " + ConceptName);
                                    DataReturn saveconcept = Utility.HandleData(this.D.SaveConcept(ConceptIdNew, ConceptName, ConceptDescription, ConceptPlayBookInfo, ConceptPlayBookClient, ConceptAllowNone));
                                    if (!saveconcept.success)
                                    {
                                        return;
                                    }

                                    ConceptIdNew = saveconcept.id;
                                    ConceptIdMapping.Add(ConceptIdOld, ConceptIdNew);

                                    first = false;
                                }

                                // create/update clauses

                                // check if we already have the clause name
                                DataReturn checkclause = Utility.HandleData(this.D.CheckClause(r["Clause__r_Name"].ToString()));
                                if (!checkclause.success)
                                {
                                    return;
                                }


                                string ClauseIdOld = r.Row["Clause__r_Id"].ToString();

                                string ClauseIdNew = "";
                                if (checkclause.dt.Rows.Count > 0)
                                {
                                    ClauseIdNew = checkclause.dt.Rows[0][0].ToString();
                                }

                                // set the clause id and update the Concept Id to the new one

                                r.Row["Clause__r_Id"] = ClauseIdNew;
                                r.Row["Clause__r_Concept__r_Id"] = ConceptIdMapping[r.Row["Clause__r_Concept__r_Id"].ToString()];

                                Globals.ThisAddIn.ProcessingUpdate((ClauseIdNew != "" ? "Updating Clause" : "Save Clause as ") + r.Row["Clause__r_Name"].ToString());

                                DataReturn saveclause = Utility.HandleData(this.D.SaveClauseFromTemplateClause(r.Row));
                                if (!saveclause.success)
                                {
                                    return;
                                }

                                ClauseIdNew = saveclause.id;
                                ClauseIdMapping.Add(ClauseIdOld, ClauseIdNew);
                                r.Row["Clause__r_Id"] = ClauseIdNew;

                                // create/update Template Clause mapping
                                string TemplateClauseIdOld = r.Row["Id"].ToString();

                                // check if we have one already - don't use name, see if there is on with the new template and new clause id
                                DataReturn checktemplateclause = Utility.HandleData(this.D.GetTemplateClause(this.Id, ClauseIdNew));
                                if (!checktemplateclause.success)
                                {
                                    return;
                                }

                                string TemplateClauseIdNew = "";
                                if (checktemplateclause.dt.Rows.Count > 0)
                                {
                                    TemplateClauseIdNew = checktemplateclause.dt.Rows[0][0].ToString();
                                }

                                string TemplateClauseName = r.Row["Name"].ToString();
                                string TemplateClauseOrder = r.Row["Order__c"].ToString();
                                string TemplateClauseDefault = r.Row["DefaultSelection__c"].ToString();


                                Globals.ThisAddIn.ProcessingUpdate((TemplateClauseIdNew != "" ? "Updating Template Clause link" : "Save Template Clause link as ") + TemplateClauseName);
                                DataReturn savetemplateclause = Utility.HandleData(this.D.SaveTemplateClause(TemplateClauseIdNew, TemplateClauseName, this.Id, ClauseIdNew, TemplateClauseOrder, TemplateClauseDefault));
                                if (!savetemplateclause.success)
                                {
                                    return;
                                }
                                r.Row["Id"] = savetemplateclause.id;

                            }
                        }
                    }


                    // now step through the Elements - copy the element and add in the caluse element mapping
                    if (this.DTElement != null)
                    {
                        List<string> OrphanIds = new List<string>();
                        foreach (DataRow r in this.DTElement.Rows)
                        {

                            // have to check that there is a matching clause
                            // reported bu 19Aug - looks like when you delete a clause the element
                            // mapping may be left in and that breaks the import cause it can't find
                            // the old clause to map to the new clause - should fix the delete TODO!
                            // but for now just check we have got the clause and if not just skip it
                            // which will actually lead to the right thing!

                            if (ClauseIdMapping.ContainsKey(r["Clause__r_Id"].ToString()))
                            {


                                // create the element - check we haven't already created the element first
                                // if so don't create it again 

                                string ElementIdOld = r["Element__r_Id"].ToString();
                                string ElementIdNew = "";
                                if (ElementIdMapping.ContainsKey(ElementIdOld))
                                {
                                    r["Element__r_Id"] = ElementIdMapping[ElementIdOld];
                                    ElementIdNew = ElementIdMapping[ElementIdOld];
                                }
                                else
                                {
                                    // check for a name match on the element

                                    DataReturn checkelement = Utility.HandleData(this.D.CheckElement(r["Element__r_Name"].ToString()));
                                    if (!checkelement.success)
                                    {
                                        return;
                                    }

                                    if (checkelement.dt.Rows.Count > 0)
                                    {
                                        ElementIdNew = checkelement.dt.Rows[0][0].ToString();
                                    }

                                    r["Element__r_Id"] = ElementIdNew;

                                    Globals.ThisAddIn.ProcessingUpdate((ElementIdNew == "" ? "Save Element as " : "Updating Element ") + r["Element__r_Name"].ToString());

                                    DataReturn saveelement = Utility.HandleData(this.D.SaveElementFromClauseElement(r));
                                    if (!saveelement.success) return;

                                    ElementIdNew = saveelement.id;
                                    ElementIdMapping.Add(ElementIdOld, ElementIdNew);
                                    r["Element__r_Id"] = ElementIdNew;
                                }

                                // create the Clause Element Entry

                                // check if we have one already - don't use name, see if there is on with the new template and new clause id

                                string ClauseElementClauseId = ClauseIdMapping[r["Clause__r_Id"].ToString()];
                                string ClauseElementName = r["Name"].ToString();
                                string ClauseElementOrder = r["Order__c"].ToString();

                                DataReturn checkclauseelement = Utility.HandleData(this.D.GetElement(ClauseElementClauseId, ElementIdNew));
                                if (!checkclauseelement.success)
                                {
                                    return;
                                }

                                string ClauseElementId = "";
                                if (checkclauseelement.dt.Rows.Count > 0)
                                {
                                    ClauseElementId = checkclauseelement.dt.Rows[0][0].ToString();
                                }

                                Globals.ThisAddIn.ProcessingUpdate((ClauseElementId == "" ? "Save Element Clause Link as " : "Updating Element Clause Link ") + ClauseElementName);

                                DataReturn saveclauseelement = Utility.HandleData(this.D.SaveClauseElement(ClauseElementId, ClauseElementName, ClauseElementClauseId, ElementIdNew, ClauseElementOrder));
                                if (!saveclauseelement.success) return;

                                // update the dataset
                                r["Id"] = saveclauseelement.id;
                                r["Clause__r_Id"] = ClauseElementClauseId;
                            }
                            else
                            {
                                // this is an orphan element row with a clause that isnt actually in the contract
                                // so remember and delete later
                                OrphanIds.Add(r["Id"].ToString());
                            }
                        }

                        // Delete out any orphans
                        for (int i = this.DTElement.Rows.Count - 1; i >= 0; i--)
                        {
                            if (OrphanIds.Contains(this.DTElement.Rows[i]["Id"].ToString())) this.DTElement.Rows[i].Delete();
                        }


                    }


                    // Now step through the Clause XML - need to update the element tags 
                    // and then update the XML

                    // switch off the save handler so we can save the clauses with the ids 
                    Globals.ThisAddIn.RemoveSaveHandler();

                    if (DTClauseXML != null)
                    {
                        int cnt = 1;
                        List<string> OrphanIds = new List<string>();
                        foreach (DataRow r in this.DTClauseXML.Rows)
                        {
                            // get the new id from the mapping - check for orpahn clauses where the concept
                            // isnt in the document - should happen but jsut incase
                            if (ClauseIdMapping.ContainsKey(r["Id"].ToString()))
                            {
                                string ClauseIdNew = ClauseIdMapping[r["Id"].ToString()];
                                string xml = r["XML"].ToString();
                                string text = "";

                                r["Id"] = ClauseIdNew;

                                // and save

                                // Save the Attachment to Salesforce
                                Globals.ThisAddIn.ProcessingUpdate("Save Clause Word File " + (cnt++).ToString());
                                string clausefilename = Utility.SaveTempFile(ClauseIdNew);

                                Word.Document scratch = Globals.ThisAddIn.Application.Documents.Add(Visible: false);

                                // get the text
                                scratch.Content.InsertXML(xml);
                                text = scratch.Range().Text;

                                // update the properties tag of the word doc and save
                                Globals.ThisAddIn.AddDocId(scratch, "ClauseTemplate", ClauseIdNew);

                                // update the Element tags
                                Globals.ThisAddIn.UpdateClauseTemplateContentControls(scratch, ConceptIdMapping, ClauseIdMapping, ElementIdMapping);

                                xml = scratch.Range().WordOpenXML;

                                scratch.SaveAs2(FileName: clausefilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                                var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                                docclosescratch.Close(false);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                                // Now save the file
                                dr = Utility.HandleData(this.D.SaveClauseFile(ClauseIdNew, text, clausefilename));
                                if (!dr.success) return;

                                // and update the dataset
                                r["XML"] = xml;
                            }
                            else
                            {
                                OrphanIds.Add(r["Id"].ToString());                                
                            }
                        }

                        // Delete out any orphans
                        for (int i = this.DTClauseXML.Rows.Count - 1; i >= 0; i--)
                        {
                            if (OrphanIds.Contains(this.DTClauseXML.Rows[i]["Id"].ToString())) this.DTClauseXML.Rows[i].Delete();
                        }
                    }


                    Globals.ThisAddIn.AddSaveHandler();

                    // now update the tempalte with the new Concept and Element Ids and save
                    // this will do the cached values for the elements in the selected clause as well

                    Globals.ThisAddIn.ProcessingUpdate("Update Template Ids");
                    Globals.ThisAddIn.UpdateContractTemplateContentControls(this.Doc, ConceptIdMapping, ClauseIdMapping, ElementIdMapping);

                    // remove the Custom Parts
                    foreach (Office.CustomXMLPart part in this.Doc.CustomXMLParts)
                    {
                        if (part.NamespaceURI == "http://www.axiomlaw.com/irisribbon")
                        {
                            part.Delete();
                        }
                    }

                    // now save the template
                    Globals.ThisAddIn.ProcessingUpdate("Save The Template");
                    this.SaveDoc();

                    // now refresh and add the content handlers
                    this.Refresh();
                    Globals.ThisAddIn.AddContentControlHandler(this.Doc);

                }

                Globals.ThisAddIn.ProcessingStop("Finished");
            }
            catch (Exception e)
            {
                string message = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) message += " " + e.InnerException.Message;
                MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }
        }

        private void btnClone_Click(object sender, RoutedEventArgs e)
        {

            // Clone! like Import but from the current instance - lots of options so need a dialog
            //
            // What are the options - Clone Template with Same Concepts and Clauses and Elements
            // Clone Template with Clones of Clauses and Clones of Concepts - need a new name

            // obvioulsy lots of other options but this will do for now!

            CloneTemplate ctemplate = new CloneTemplate();
            ctemplate.Open(this);
            ctemplate.Show();


        }

        public void DoClone(string mode, string newname, string prependname)
        {
            // MessageBox.Show("Get to it! " + mode + " | " + newname + "|" + prependname);

            // currently mode is either CloneTemplate OR CloneTemplateConceptClause
            // if the latter then the prependname should be used to copy concepts and clauses
            try
            {
                Globals.ThisAddIn.ProcessingStart("Clone Template");

               
                // ok this is just like an import but instead of stepping through 
                // the data in the exported template just step through the current one
                // now we need to save everything and update the ids

                // keep list of the old ids
                string OldTemplateId = this.Id;
                Dictionary<string, string> ConceptIdMapping = new Dictionary<string, string>();
                Dictionary<string, string> ClauseIdMapping = new Dictionary<string, string>();
                Dictionary<string, string> ElementIdMapping = new Dictionary<string, string>();

                // first the template
                string newtemplateid = "";
                string newtemplatename = newname;

                // look for a Template with a matching name
                DataReturn checktemplate = Utility.HandleData(this.D.CheckTemplate(newname));
                if (!checktemplate.success)
                {
                    return;
                }

                // check if we have a match and if so warn them
                bool updatetemplate = false;
                if (checktemplate.dt.Rows.Count > 0)
                {
                    MessageBoxResult res = MessageBox.Show("There is already a template called '" + newname + "'  Would you like to update from this template?", "Warning", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.Cancel)
                    {
                        Globals.ThisAddIn.ProcessingStop("Finished");
                        // hide the sidebar
                        Globals.ThisAddIn.ShowTaskPane(false);
                        return;
                    }
                    newtemplateid = checktemplate.dt.Rows[0][0].ToString();
                    updatetemplate = true;
                }

                this.Id = newtemplateid;
                this.DTTemplate.Rows[0]["Id"] = newtemplateid;
                this.DTTemplate.Rows[0]["Name"] = newtemplatename;

                if (newtemplateid == "") Globals.ThisAddIn.ProcessingUpdate("Create New Template as " + newtemplatename);

                DataReturn dr = Utility.HandleData(this.D.SaveTemplate(this.DTTemplate.Rows[0]));
                if (!dr.success)
                {
                    MessageBox.Show("There has been an issue - template will not be cloned");
                    return;
                }

                // remember the new id            
                this.Id = dr.id;
                this.DTTemplate.Rows[0]["Id"] = dr.id;
                this.DTTemplate.Rows[0]["Name"] = newtemplatename;


                // set the template id in the doc to get things to work
                Globals.ThisAddIn.AddDocId(this.Doc, "ContractTemplate", this.Id);

                // update the form
                Utility.UpdateForm(new Grid[] { formGridTemplate }, ((DataRowView)this.DTTemplate.DefaultView[0]).Row);
                Utility.ReadOnlyForm(false, new Grid[] { formGridTemplate });
                this.btnSave.IsEnabled = false;
                this.btnCancel.IsEnabled = false;

                // now step through the clauses and create any that aren't there
                // step through in the order of the concepts in the template
                Globals.ThisAddIn.ProcessingUpdate("Get Concept Order");
                string conceptorder = Globals.ThisAddIn.GetConceptOrder(this.Doc);

                if (conceptorder != "")
                {
                    string[] cotags = conceptorder.Split(',');

                    foreach (string cotag in cotags)
                    {

                        string[] conceptdetails = cotag.Split('|');

                        // format is Concept|ConceptId|ClauseId|LastModified - last 2 may not be there
                        string concept = conceptdetails[1];

                        //Populate all the Concepts and Clauses                            
                        DataView dv = new DataView(this.DTClause);
                        dv.RowFilter = "Clause__r_Concept__r_Id='" + concept + "'";
                        dv.Sort = "Order__c";

                        bool first = true;
                        string ConceptName = "";

                        foreach (DataRowView r in dv)
                        {
                            if (first)
                            {

                                // if mode is CloneTemplate then don't need to create or update just use the id as is
                                // if mode is CloneTemplateConceptClause - then add prepend to the name
                                string conceptname = r["Clause__r_Concept__r_Name"].ToString();
                                string ConceptIdOld = r["Clause__r_Concept__r_Id"].ToString();
                                string ConceptIdNew = "";

                                if (mode == "CloneTemplate")
                                {
                                    // just use Concept as is                                
                                    ConceptIdNew = ConceptIdOld;
                                    ConceptIdMapping.Add(ConceptIdOld, ConceptIdNew);

                                }
                                else if (mode == "CloneTemplateConceptClause")
                                {
                                    // prepend name and check if its there
                                    conceptname = Utility.Truncate(prependname + conceptname, 80);

                                    // check if we already have the concept name
                                    DataReturn checkconcept = Utility.HandleData(this.D.CheckConcept(conceptname));
                                    if (!checkconcept.success)
                                    {
                                        return;
                                    }

                                    if (checkconcept.dt.Rows.Count > 0)
                                    {
                                        ConceptIdNew = checkconcept.dt.Rows[0][0].ToString();
                                    }

                                    ConceptName = conceptname;
                                    string ConceptDescription = r["Clause__r_Concept__r_Description__c"].ToString();
                                    string ConceptPlayBookInfo = r["Clause__r_Concept__r_PlayBookInfo__c"].ToString();
                                    string ConceptPlayBookClient = r["Clause__r_Concept__r_PlayBookClient__c"].ToString();
                                    bool ConceptAllowNone = Convert.ToBoolean(r["Clause__r_Concept__r_AllowNone__c"]);


                                    Globals.ThisAddIn.ProcessingUpdate(ConceptIdNew != "" ? "Updating Concept " + ConceptName : "Save Concept as " + ConceptName);
                                    DataReturn saveconcept = Utility.HandleData(this.D.SaveConcept(ConceptIdNew, ConceptName, ConceptDescription, ConceptPlayBookInfo, ConceptPlayBookClient, ConceptAllowNone));
                                    if (!saveconcept.success)
                                    {
                                        return;
                                    }

                                    ConceptIdNew = saveconcept.id;

                                    // shouldn't already be in the bag bust just to be safe
                                    if (!ConceptIdMapping.ContainsKey(ConceptIdOld))
                                    {
                                        ConceptIdMapping.Add(ConceptIdOld, ConceptIdNew);
                                    }
                                }

                                first = false;
                            }

                            // create/update clauses
                            // if mode is CloneTemplate then don't need to create or update just use the id as is
                            // if mode is CloneTemplateConceptClause - then add prepend to the name

                            string clausename = r["Clause__r_Name"].ToString();
                            string ClauseIdOld = r.Row["Clause__r_Id"].ToString();
                            string ClauseIdNew = "";

                            if (mode == "CloneTemplate")
                            {
                                // just use Clause as is
                                ClauseIdNew = ClauseIdOld;
                                ClauseIdMapping.Add(ClauseIdOld, ClauseIdNew);
                            }
                            else if (mode == "CloneTemplateConceptClause")
                            {
                                // prepend name and check if its there
                                clausename = Utility.Truncate(prependname + clausename, 80);
                                // check if we already have the clause name
                                DataReturn checkclause = Utility.HandleData(this.D.CheckClause(clausename));
                                if (!checkclause.success)
                                {
                                    return;
                                }

                                if (checkclause.dt.Rows.Count > 0)
                                {
                                    ClauseIdNew = checkclause.dt.Rows[0][0].ToString();
                                }
                                
                                // set the clause id and update the Concept Id to the new one

                                r.Row["Clause__r_Id"] = ClauseIdNew;
                                r.Row["Clause__r_Concept__r_Id"] = ConceptIdMapping[r.Row["Clause__r_Concept__r_Id"].ToString()];
                                r.Row["Clause__r_Concept__r_Name"] = ConceptName;
                                r.Row["Clause__r_Name"] = clausename;

                                Globals.ThisAddIn.ProcessingUpdate((ClauseIdNew != "" ? "Updating Clause" : "Save Clause as ") + clausename);

                                DataReturn saveclause = Utility.HandleData(this.D.SaveClauseFromTemplateClause(r.Row));
                                if (!saveclause.success)
                                {
                                    return;
                                }

                                ClauseIdNew = saveclause.id;

                                // need to check that its not in the bag already - this could happen if the template
                                // has two clauses with the same name - shouldn't really happen but hey ho
                                if (!ClauseIdMapping.ContainsKey(ClauseIdOld))
                                {
                                    ClauseIdMapping.Add(ClauseIdOld, ClauseIdNew);
                                }
                                r.Row["Clause__r_Id"] = ClauseIdNew;
                            }





                            // create/update Template Clause mapping
                            string TemplateClauseIdOld = r.Row["Id"].ToString();

                            // check if we have one already - don't use name, see if there is on with the new template and new clause id
                            DataReturn checktemplateclause = Utility.HandleData(this.D.GetTemplateClause(this.Id, ClauseIdNew));
                            if (!checktemplateclause.success)
                            {
                                return;
                            }

                            string TemplateClauseIdNew = "";
                            if (checktemplateclause.dt.Rows.Count > 0)
                            {
                                TemplateClauseIdNew = checktemplateclause.dt.Rows[0][0].ToString();
                            }

                            string TemplateClauseName = Utility.Truncate(newtemplatename, 35) + "-" + Utility.Truncate(clausename, 35);
                            string TemplateClauseOrder = r.Row["Order__c"].ToString();
                            string TemplateClauseDefault = r.Row["DefaultSelection__c"].ToString();


                            Globals.ThisAddIn.ProcessingUpdate((TemplateClauseIdNew != "" ? "Updating Template Clause link" : "Save Template Clause link as ") + TemplateClauseName);
                            DataReturn savetemplateclause = Utility.HandleData(this.D.SaveTemplateClause(TemplateClauseIdNew, TemplateClauseName, this.Id, ClauseIdNew, TemplateClauseOrder, TemplateClauseDefault));
                            if (!savetemplateclause.success)
                            {
                                return;
                            }

                            // check this isn't a duplicate mapping - if there were 2 mapping records with the same template/clause then this will fail because 
                            // the second one will get the first id back and that will already be in the table

                            if(this.DTClause.Select("Id='" + savetemplateclause.id + "'").Length==0){
                                r.Row["Id"] = savetemplateclause.id;
                            }
                            else
                            {
                                r.Delete();
                            }

                        }
                    }
                }


                // now step through the Elements
                // we are using the same elements so don't need to create those
                // only need to upodate the link if we are cloning the clauses

                if (mode == "CloneTemplateConceptClause")
                {
                    if (this.DTElement != null)
                    {

                        List<string> OrphanIds = new List<string>();
                        foreach (DataRow r in this.DTElement.Rows)
                        {

                            // have to check that there is a matching clause
                            // reported bu 19Aug - looks like when you delete a clause the element
                            // mapping may be left in and that breaks the import cause it can't find
                            // the old clause to map to the new clause - should fix the delete TODO!
                            // but for now just check we have got the clause and if not just skip it
                            // which will actually lead to the right thing!

                            if (ClauseIdMapping.ContainsKey(r["Clause__r_Id"].ToString()))
                            {


                                string ElementId = r["Element__r_Id"].ToString();

                                // create the Clause Element Entry

                                // check if we have one already - don't use name, see if there is on with the new template and new clause id

                                string ClauseElementClauseId = ClauseIdMapping[r["Clause__r_Id"].ToString()];
                                string ClauseElementName = Utility.Truncate(r["Clause__r_Name"].ToString(), 35) + "-" + Utility.Truncate(r["Element__r_Name"].ToString(), 35);
                                string ClauseElementOrder = r["Order__c"].ToString();

                                DataReturn checkclauseelement = Utility.HandleData(this.D.GetElement(ClauseElementClauseId, ElementId));
                                if (!checkclauseelement.success)
                                {
                                    return;
                                }

                                string ClauseElementId = "";
                                if (checkclauseelement.dt.Rows.Count > 0)
                                {
                                    ClauseElementId = checkclauseelement.dt.Rows[0][0].ToString();
                                }

                                Globals.ThisAddIn.ProcessingUpdate((ClauseElementId == "" ? "Save Element Clause Link as " : "Updating Element Clause Link ") + ClauseElementName);

                                DataReturn saveclauseelement = Utility.HandleData(this.D.SaveClauseElement(ClauseElementId, ClauseElementName, ClauseElementClauseId, ElementId, ClauseElementOrder));
                                if (!saveclauseelement.success) return;

                                // update the dataset
                                r["Id"] = saveclauseelement.id;
                                r["Clause__r_Id"] = ClauseElementClauseId;
                            }
                            else
                            {
                                // this is an orphan element row with a clause that isnt actually in the contract
                                // so jsut remember the id and delete it out after
                                OrphanIds.Add(r["Id"].ToString());                                 
                            }
                        }

                        // Delete out any orphans
                        for (int i = this.DTElement.Rows.Count - 1; i >= 0; i--)
                        {
                            if (OrphanIds.Contains(this.DTElement.Rows[i]["Id"].ToString())) this.DTElement.Rows[i].Delete();
                        }
                    }
                }


                // Now step through the Clause XML - need to update the element tags 
                // and then update the XML

                // switch off the save handler so we can save the clauses with the ids 
                Globals.ThisAddIn.RemoveSaveHandler();

                // only have to do this if we are cloning the clauses
                if (mode == "CloneTemplateConceptClause")
                {
                    if (DTClauseXML != null)
                    {
                        int cnt = 1;
                        List<string> OrphanIds = new List<string>();
                        foreach (DataRow r in this.DTClauseXML.Rows)
                        {
                            // get the new id from the mapping - check for orpahn clauses where the concept
                            // isnt in the document - should happen but jsut incase
                            if (ClauseIdMapping.ContainsKey(r["Id"].ToString()))
                            {
                                string ClauseIdNew = ClauseIdMapping[r["Id"].ToString()];
                                string xml = r["XML"].ToString();
                                string text = "";

                                r["Id"] = ClauseIdNew;

                                // and save

                                // Save the Attachment to Salesforce
                                Globals.ThisAddIn.ProcessingUpdate("Save Clause Word File " + (cnt++).ToString());
                                string clausefilename = Utility.SaveTempFile(ClauseIdNew);

                                Word.Document scratch = Globals.ThisAddIn.Application.Documents.Add(Visible: false);

                                // get the text
                                scratch.Content.InsertXML(xml);
                                text = scratch.Range().Text;

                                // update the properties tag of the word doc and save
                                Globals.ThisAddIn.AddDocId(scratch, "ClauseTemplate", ClauseIdNew);

                                // update the Element tags
                                Globals.ThisAddIn.UpdateClauseTemplateContentControls(scratch, ConceptIdMapping, ClauseIdMapping, ElementIdMapping);

                                xml = scratch.Range().WordOpenXML;

                                scratch.SaveAs2(FileName: clausefilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                                var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                                docclosescratch.Close(false);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                                // Now save the file
                                dr = Utility.HandleData(this.D.SaveClauseFile(ClauseIdNew, text, clausefilename));
                                if (!dr.success) return;

                                // and update the dataset
                                r["XML"] = xml;
                            }
                            else
                            {
                                OrphanIds.Add(r["Id"].ToString());
                            }
                        }

                        // Delete out any orphans
                        for (int i = this.DTClauseXML.Rows.Count - 1; i >= 0; i--)
                        {
                            if (OrphanIds.Contains(this.DTClauseXML.Rows[i]["Id"].ToString())) this.DTClauseXML.Rows[i].Delete();
                        }
                    }
                }


                Globals.ThisAddIn.AddSaveHandler();

                // now update the tempalte with the new Concept and Element Ids and save
                // this will do the cached values for the elements in the selected clause as well

                Globals.ThisAddIn.ProcessingUpdate("Update Template Ids");
                Globals.ThisAddIn.UpdateContractTemplateContentControls(this.Doc, ConceptIdMapping, ClauseIdMapping, ElementIdMapping);

                // now save the template
                Globals.ThisAddIn.ProcessingUpdate("Save The Template");
                this.SaveDoc();

                // now refresh and add the content handlers
                this.Refresh();
                Globals.ThisAddIn.AddContentControlHandler(this.Doc);

                Globals.ThisAddIn.ProcessingStop("Finished");

            }
            catch (Exception e)
            {
                string message = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) message += " " + e.InnerException.Message;
                MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }

        }

        private void btnDebug1_Click(object sender, RoutedEventArgs e)
        {
            // Unlock all Cluases - used for trick Section issues
            object start = this.Doc.Content.Start;
            object end = this.Doc.Content.End;
            Word.Range r = this.Doc.Range(ref start, ref end);
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                string tag = cc.Tag;
                if (tag.Contains("Concept"))
                {
                    cc.LockContents = false;
                }
            }

        }

        private void btnDebug2_Click(object sender, RoutedEventArgs e)
        {
            // Lock all Cluases - used for trick Section issues
            object start = this.Doc.Content.Start;
            object end = this.Doc.Content.End;
            Word.Range r = this.Doc.Range(ref start, ref end);
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                string tag = cc.Tag;
                if (tag.Contains("Concept"))
                {
                    cc.LockContents = true;
                }
            }
        }

        private void btnDebug3_Click(object sender, RoutedEventArgs e)
        {

            // Export out the playbook data - we've had a cipher issue that has scramble it 
            // so want to be able to update - create a word doc with the data stored as XML
            // just need name and Playbook field - will update by matching to name

            Globals.ThisAddIn.ProcessingUpdate("Export Playbook Fields");
            Word.Document export = Globals.ThisAddIn.Application.Documents.Add();

            Globals.ThisAddIn.AddDocId(export, "ExportPlaybookFields", "123456789");

            DataReturn rt = Utility.HandleData(this.D.GetConceptsPlaybookInfo());
            if (!rt.success)
            {
                return;
            }

            // now get the meta data and store it as custom xml parts
            DataSet ds = new DataSet();
            ds.Namespace = "http://www.axiomlaw.com/irisribbon";
            rt.dt.TableName = "ConceptPlaybookData";
            ds.Tables.Add(rt.dt);

            string xmldata = "";
            using (StringWriter stringWriter = new StringWriter())
            {
                ds.WriteXml(new XmlTextWriter(stringWriter));
                xmldata = stringWriter.ToString();
            };

            Office.CustomXMLPart data = export.CustomXMLParts.Add(xmldata);

            export.Activate();

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Word Document (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docx";
            dlg.FileName = "ExportData-Playbook";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                export.SaveAs2(dlg.FileName);
            }
        }

        private void btnDebug4_Click(object sender, RoutedEventArgs e)
        {


            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Word Document (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docx";
            Nullable<bool> result = dlg.ShowDialog();
            Word.Document import;
            if (result == true)
            {
                import = Globals.ThisAddIn.Application.Documents.Open(dlg.FileName);
            }
            else
            {
                return;
            }


            Globals.ThisAddIn.ProcessingStart("Import Data");

            Office.CustomXMLPart data = null;
            foreach (Office.CustomXMLPart part in import.CustomXMLParts)
            {
                if (part.NamespaceURI == "http://www.axiomlaw.com/irisribbon")
                {
                    data = part;
                }
            }

            if (data != null)
            {

                Globals.ThisAddIn.ProcessingStart("Read Data from Document");

                DataSet ds = new DataSet();

                System.IO.StringReader xmlSR = new System.IO.StringReader(data.DocumentElement.XML);

                ds.ReadXml(xmlSR, XmlReadMode.Auto);

                DataTable dr = ds.Tables["ConceptPlaybookData"];
                foreach (DataRow r in dr.Rows)
                {

                    string ClauseName = r["Name"].ToString();
                    string PlaybookInfo = r["PlayBookInfo__c"].ToString();
                    string PlaybookClient = r["PlayBookClient__c"].ToString();

                    if (ClauseName != "" && (PlaybookInfo != "" || PlaybookClient != ""))
                    {

                        // update if there is a match
                        DataReturn checkconcept = Utility.HandleData(this.D.CheckConcept(ClauseName));
                        if (!checkconcept.success)
                        {
                            return;
                        }

                        string ConceptIdNew = "";
                        if (checkconcept.dt.Rows.Count > 0)
                        {

                            Globals.ThisAddIn.ProcessingUpdate(" Updating Clause " + ClauseName + "");
                            ConceptIdNew = checkconcept.dt.Rows[0][0].ToString();
                            DataReturn saveclause = Utility.HandleData(this.D.SaveConcept(ConceptIdNew, PlaybookInfo, PlaybookClient));



                        }
                        else
                        {
                            Globals.ThisAddIn.ProcessingUpdate("No Match for Clause " + ClauseName);
                        }

                    }
                    else
                    {
                        Globals.ThisAddIn.ProcessingUpdate("Clause " + ClauseName + " has no playbook info");

                    }

                }

                Globals.ThisAddIn.ProcessingStop("Finished");
            }
        }


        public void DoCleanUp(bool dostuff)
        {

            // Look through the template and look for problems
            // 1 - Orphans
            // 2 - Concepts or Clauses with duplicate names


            string message = "";
            bool issue = false;

            try
            {
                Globals.ThisAddIn.ProcessingStart("Clean Up");

                                
                // ORPHANS - document is the master - check that the database doesn't have clauses or elements that 
                // are not in the document

                Globals.ThisAddIn.ProcessingUpdate("Orphans");
                string conceptorder = Globals.ThisAddIn.GetConceptOrder(this.Doc);
                string orphanmessage = "";
                int orphancount = 0;
                List<string> orphanids = new List<string>();
                if (conceptorder != "")
                {
                    string[] cotags = conceptorder.Split(',');

                    foreach (DataRow r in this.DTClause.Rows)
                    {
                        string ConceptName = r["Clause__r_Concept__r_Name"].ToString();
                        string ConceptId = r["Clause__r_Concept__r_Id"].ToString();
                        string ClauseName = r["Clause__r_Name"].ToString();
                        string ClauseId = r["Clause__r_Id"].ToString();
                        string Id = r["Id"].ToString();
                        string Name = r["Name"].ToString();

                        // check there is a matching concept in the document
                        bool match = false;

                        foreach (string cotag in cotags)
                        {
                            // format is Concept|ConceptId|ClauseId|LastModified - last 2 may not be there
                            string[] conceptdetails = cotag.Split('|');
                            if (conceptdetails.Length > 1)
                            {
                                if (conceptdetails[1] == ConceptId) match = true;
                            }
                            else
                            {
                                // bad tag
                            }
                        }

                        if (!match)
                        {
                            orphanids.Add(Id);
                            orphancount++;
                            orphanmessage += "\tOrphan Link: " + Name + "\n"; // + " Clause:" + ClauseName + " Concept:" + ConceptName + "\n";
                        }
                    }                    
                }

                if (orphancount > 0)
                {
                    message += "1a. " +  orphancount + " Orphan Clause Links\n";
                    message += orphanmessage;
                    issue = true;
                } else {
                    message += "1a. There are NO Orphan Clause Links\n";
                }

                if (dostuff)
                {
                    // step through list and do the delete
                    foreach (string templateclauseid in orphanids)
                    {
                        Globals.ThisAddIn.ProcessingUpdate(" Delete Template Clause: " + templateclauseid + "");
                        // delete out the orphans
                        message += "\t" + " Delete Template Clause: " + templateclauseid + "\n";
                        DataReturn dr = Utility.HandleData(this.D.DeleteTemplateClause(templateclauseid));
                        if (!dr.success) return;

                        // delete from cache
                        DataRow deleterow = null;
                        foreach (DataRow r in this.DTClause.Rows)
                        {
                            if (r["Id"].ToString() == templateclauseid)
                            {
                                deleterow = r;
                            }
                        }
                        if (deleterow != null) deleterow.Delete();
                    }
                    // need to reload the elements
                }


                // Orphan elements - the orphan elements issue is actually because of the orphan clauses
                // the elements belong to the clause so when they get orphaned we get element orphans
                // there is the issue of elements that are in the database but not in the template but come
                // back to that when we do more element work


                // Orphan Content Controls - check there are no controls in the doc that don't match values in the database

                // First Clause - Check the Clause is there and Check that the select Clause exists
                if (conceptorder != "")
                {
                    List<string> orphanconceptcontrols = new List<string>();
                    List<string> baddefaultclause = new List<string>();
                    string[] cotags = conceptorder.Split(',');

                    foreach (string cotag in cotags)
                    {
                        // format is Concept|ConceptId|ClauseId|LastModified - last 2 may not be there
                        string[] conceptdetails = cotag.Split('|');

                        string conceptid = "";
                        if (conceptdetails.Length > 1) conceptid = conceptdetails[1];

                        string clauseid = "";
                        if (conceptdetails.Length > 2) clauseid = conceptdetails[2];

                        bool matchconcept = false;
                        bool matchclause = false;

                        foreach (DataRow r in this.DTClause.Rows)
                        {
                            // Check Concept exists                                
                            if (conceptid == r["Clause__r_Concept__r_Id"].ToString()) matchconcept = true;

                            // Check Clause exist
                            if (clauseid != "")
                            {
                                if (clauseid == r["Clause__r_Id"].ToString()) matchclause = true;                                
                            }
                        }

                        if (!matchconcept)
                        {
                            if (!orphanconceptcontrols.Contains(conceptid)) orphanconceptcontrols.Add(conceptid);
                        }

                        if (!matchclause && clauseid!="")
                        {
                            if (!baddefaultclause.Contains(conceptid)) baddefaultclause.Add(conceptid);
                        }
                    }

                    if (orphanconceptcontrols.Count > 0)
                    {
                        message += "\n1b. " + orphanconceptcontrols.Count + " Orphan Clause Controls in doc\n";                        
                        issue = true;
                    }
                    else
                    {
                        message += "\n1b. There are NO Orphan Clause Controls in doc\n";
                    }

                    if (baddefaultclause.Count > 0)
                    {
                        message += "\n1c. " + baddefaultclause.Count + " Bad default Clause selections\n";
                        // get the name of the clause
                        foreach (string id in baddefaultclause)
                        {
                            foreach (DataRow r in this.DTClause.Rows)
                            {
                                if (r["Clause__r_Concept__r_Id"].ToString() == id)
                                {
                                    message += "\t Concept:" + r["Clause__r_Concept__r_Name"].ToString() + "\n";
                                    break;
                                }
                            }
                        }
                        issue = true;
                    }
                    else
                    {
                        message += "\n1c. There are NO Bad default Clause selections\n";
                    }

                    if (dostuff)
                    {
                        if (orphanconceptcontrols.Count > 0)
                        {
                            // Remove the Clause Controls that don't have database links - just remove the control
                            foreach(string id in orphanconceptcontrols){
                                Globals.ThisAddIn.ProcessingUpdate(" REMOVE Control " + id);
                                message += "\t REMOVE Control " + id;
                                Globals.ThisAddIn.RemoveConcept(this.Doc, id);
                            }
                        }

                        if (baddefaultclause.Count > 0)
                        {
                            // Remove the default clause indication from the control
                            foreach (string id in baddefaultclause)
                            {
                                Globals.ThisAddIn.ProcessingUpdate(" take out default clause indicator " + id);
                                message += "\t take out default clause indicator " + id;
                                Globals.ThisAddIn.UpdateContractTemplatesConceptTag(this.Doc, id,"","");
                            }
                        }

                    }
                }







               
                // DUPLICATE NAMES
                Globals.ThisAddIn.ProcessingUpdate("Duplicates");
                Dictionary<string, string> conceptnames = new Dictionary<string, string>();
                Dictionary<string,int> conceptdupes = new Dictionary<string,int>();                                             
                Dictionary<string, string> clausenames = new Dictionary<string, string>();
                Dictionary<string, int> clausedupes = new Dictionary<string, int>();

                foreach (DataRow r in this.DTClause.Rows)
                {
                    string ConceptName = r["Clause__r_Concept__r_Name"].ToString();
                    string ConceptId = r["Clause__r_Concept__r_Id"].ToString();

                    if (conceptnames.ContainsKey(ConceptId))
                    {
                        // we arleady have this id so this is ok
                    }
                    else
                    {
                        // check if we have the name
                        if (conceptnames.ContainsValue(ConceptName))
                        {
                            // diferent Id and Name!
                            if(conceptdupes.ContainsKey(ConceptName)){
                                conceptdupes[ConceptName] += 1;
                            } else {
                                conceptdupes.Add(ConceptName,1);
                            }

                            // add it in
                            conceptnames.Add(ConceptId, ConceptName);

                        }
                        else
                        {
                            // add it in
                            conceptnames.Add(ConceptId, ConceptName);
                        }
                    }

                    string ClauseName = r["Clause__r_Name"].ToString();
                    string ClauseId = r["Clause__r_Id"].ToString();

                    if (clausenames.ContainsKey(ClauseId))
                    {
                        // we arleady have this id so this is ok
                    }
                    else
                    {
                        // check if we have the name
                        if (clausenames.ContainsValue(ClauseName))
                        {
                            // diferent Id and Name!
                            if (clausedupes.ContainsKey(ClauseName))
                            {
                                clausedupes[ClauseName] += 1;
                            }
                            else
                            {
                                clausedupes.Add(ClauseName, 1);
                            }
                            // add it anyway so we have it in the list for the dostuff
                            clausenames.Add(ClauseId, ClauseName);
                        }
                        else
                        {
                            // add it in
                            clausenames.Add(ClauseId, ClauseName);
                        }
                    }
                }


                if (conceptdupes.Count > 0)
                {
                    issue = true;
                    message += "\n2. " + conceptdupes.Count + " Duplicate Concept Names \n";
                    foreach(string key in conceptdupes.Keys){
                        message += "\t" + key + " occurs " + (conceptdupes[key]+1).ToString() + " times\n";
                    }
                }
                else
                {
                    message += "\n2. There are NO Duplicate Concept Names\n";
                }


                if (dostuff)
                {
                    // step through the duplicate concepts and rename them!
                    
                    foreach (string key in conceptdupes.Keys)
                    {
                        List<string> matchingids = new List<string>();
                        // message += "\t" + key + " occurs " + (conceptdupes[key] + 1).ToString() + " times\n";
                        string conceptname = key;
                        // get the ids                        
                        foreach (string id in conceptnames.Keys)
                        {
                            if (conceptnames[id] == conceptname) matchingids.Add(id);
                        }


                        // ok now have a List of Ids rename the 2nd, 3rd etc. - skip the first one
                        for (int i = 1; i < matchingids.Count; i++)
                        {
                            string conceptid = matchingids[i];
                            string newconceptname = Utility.Truncate(conceptname,77) + "-" + i.ToString();

                            // rename - make sure we don't match one in the template
                            // doesn't matter if we match in another template

                            bool clash = false;
                            int z = 1;
                            do
                            {
                                foreach (DataRow r in this.DTClause.Rows)
                                {
                                    if (r["Clause__r_Concept__r_Name"].ToString() == newconceptname)
                                    {
                                        clash = true;
                                    }
                                }
                                if (clash) newconceptname += "-" + z.ToString();
                                z++;
                            } while (clash || z > 10);

                            if (z > 10)
                            {
                                message += "\t Problem renaming Concept: " + conceptname + " too many clashes";
                            }
                            else
                            {
                                message += "\tRename Concept: " + conceptname + " to " + newconceptname + "\n";

                                // do in cache table
                                foreach (DataRow r in this.DTClause.Rows)
                                {
                                    if (r["Clause__r_Concept__r_Id"].ToString() == conceptid)
                                    {
                                        r["Clause__r_Concept__r_Name"] = newconceptname;
                                    }
                                }

                                // update sforce
                                Globals.ThisAddIn.ProcessingUpdate(" Update Concept Name" + conceptname + " to " + newconceptname);
                                Utility.HandleData(this.D.UpdateConceptName(conceptid, newconceptname));
                            }

                        }
                    }


                }


                               
                if (clausedupes.Count > 0)
                {
                    issue = true;
                    message += "\n3. " + clausedupes.Count + " Duplicate Clause Names \n";
                    foreach (string key in clausedupes.Keys)
                    {
                        message += "\t" + key + " occurs " + (clausedupes[key]+1).ToString() + " times\n";
                    }
                }
                else
                {
                    message += "3. There are NO Duplicate Clause Names\n";
                }

                if (dostuff)
                {
                    // step through the duplicate clauses and rename them!

                    foreach (string key in clausedupes.Keys)
                    {
                        // message += "\t" + key + " occurs " + clausedupes[key].ToString() + " times\n";
                        List<string> matchingids = new List<string>();
                        string clausename = key;
                        // get the ids                        
                        foreach (string id in clausenames.Keys)
                        {
                            if (clausenames[id] == clausename) matchingids.Add(id);
                        }

                        // ok now have a List of Ids rename the 2nd, 3rd etc. - skip the first one
                        for (int i = 1; i < matchingids.Count; i++)
                        {
                            string clauseid = matchingids[i];
                            string newclausename = Utility.Truncate(clausename,77) + "-" + i.ToString();

                            // rename - make sure we don't match one in the template
                            // doesn't matter if we match in another template

                            bool clash = false;
                            int z = 1;
                            do
                            {
                                foreach (DataRow r in this.DTClause.Rows)
                                {
                                    if (r["Clause__r_Name"].ToString() == newclausename)
                                    {
                                        clash = true;
                                    }
                                }
                                if (clash) newclausename += "-" + z.ToString();
                                z++;
                            } while (clash || z > 10);

                            if (z > 10)
                            {
                                message += "\t Problem renaming Clause: " + clausename + " too many clashes";
                            }
                            else
                            {
                                message += "\tRename Clause: " + clausename + " to " + newclausename + "\n";

                                // do in cache table
                                foreach (DataRow r in this.DTClause.Rows)
                                {
                                    if (r["Clause__r_Id"].ToString() == clauseid)
                                    {
                                        r["Clause__r_Name"] = newclausename;
                                    }
                                }

                                // update sforce
                                Globals.ThisAddIn.ProcessingUpdate(" Update Clause Name" + clausename + " to " + newclausename);
                                Utility.HandleData(this.D.UpdateClauseName(clauseid, newclausename));
                            }

                        }
                    }
                }


                // DUPLICATE template clause id entries - only require one
                List<string> clauseids = new List<string>();
                List<string> templateclauseiddupes = new List<string>();
                string dupemessage = "";
                foreach (DataRow r in this.DTClause.Rows)
                {
                    string Id = r["Id"].ToString();
                    string ClauseId = r["Clause__r_Id"].ToString();
                    string ClauseName = r["Clause__r_Name"].ToString();

                    if(clauseids.Contains(ClauseId)){
                        // shouldn't be here twice!
                        templateclauseiddupes.Add(Id);
                        dupemessage += "\t Clause " + ClauseName + " has a duplicate mapping record";
                    }
                    else
                    {
                        clauseids.Add(ClauseId);
                    }
                }

                if (templateclauseiddupes.Count > 0)
                {
                    message += "4. " + templateclauseiddupes.Count  + " Duplicate Template Clause records\n";
                    message += dupemessage;
                }
                else
                {
                    message += "4. There are NO Duplicate Template Clause records\n";
                }

                if (dostuff)
                {
                    foreach (DataRow r in this.DTClause.Rows)
                    {
                        if (templateclauseiddupes.Contains(r["Id"].ToString()))
                        {
                            Utility.HandleData(this.D.DeleteTemplateClause(r["Id"].ToString()));
                            r.Delete();                            
                        }
                    }
                }




                if (dostuff)
                {
                    this.Refresh();
                }

                Globals.ThisAddIn.ProcessingStop("Finished");

                if (!dostuff && issue)
                {
                    message += "\n\n ** HIT YES TO FIX ISSUES **";
                    var result = MessageBox.Show(message, "CleanUp", MessageBoxButton.YesNo);
                    if (result == MessageBoxResult.Yes)
                    {
                        this.DoCleanUp(true);
                    }

                }
                else
                {
                    MessageBox.Show(message);
                }
                

                
            }
            catch (Exception e)
            {
                string errormessage = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) errormessage += " " + e.InnerException.Message;

                // add in the message incase that is any help
                errormessage  += "\n\n" + message;

                MessageBox.Show(errormessage);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }
        }

        private void btnCleanUp_Click(object sender, RoutedEventArgs e)
        {
            this.DoCleanUp(false);
        }




    }
}
