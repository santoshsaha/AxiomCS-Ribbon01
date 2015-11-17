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
using Word = Microsoft.Office.Interop.Word;
using Telerik.Windows.Controls;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;
using System.Windows.Markup;
using System.Diagnostics;
using HTMLConverter;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;

namespace AxiomIRISRibbon.ContractEdit
{
    /// <summary>
    /// Interaction logic for SForceEditSideBar.xaml
    /// </summary>
    public partial class SForceEditSideBar2 : UserControl
    {

        // **TODO** - big time technical debt here - just merging the two sidebars
        // NEED TO REWRITE THE CLAUSE CHOICE to use telerik controls and work more like
        // the template edit with the cached datarecords but in a hurry for the release so for mow just
        // merge

        private Data _d;
        private SForceEdit.SObjectDef _sDocumentObjectDef;
        private SForceEdit.SObjectDef _sMatterObjectDef;
        private SForceEdit.SObjectDef _sRequestObjectDef;
        private SForceEdit.SObjectDef _sActivityObjectDef; // to do add activites to the sidebar
        private System.Windows.Media.Color _gbborder;
        private bool _setgbborder = false;
        private string _parentType;
        private string _parentId;
        private string _filename;

        private string _attachmentid;

        private DataRow _DocumentRow;
        private DataRow _MatterRow;
        private DataRow _RequestRow;

        private bool _DocumentChanges;
        private bool _MatterChanges;
        private bool _RequestChanges;

        private BackgroundWorker _saveBackgroundWorker;
        private BackgroundWorker _cloneBackgroundWorker;

        /* --- clause choice stuff ---*/
        private Word.Document _doc;

        private string _templateid;
        private string _templateplaybooklink;

        private string _matterid;
        private string _versionid;
        private string _versionclonename;
        private string _versionclonenumber;
        private string _versionclonenewdocpath;
        private bool _versioncloneattachedmode;

        private bool _attachedmode;

        //a dictionary with pointer to the clauses 
        private Dictionary<string, FrameworkElement> _clauses;

        //a dictionary with pointers to the element controls so we don't have to travers the UI tree to get them
        private Dictionary<string, FrameworkElement> _elements;


        public SForceEditSideBar2()
        {
            InitializeComponent();
            Utility.setTheme(this);

            if (StyleManager.ApplicationTheme.ToString() == "Windows8" || StyleManager.ApplicationTheme.ToString() == "Expression_Dark")
            {
                _setgbborder = true;
                _gbborder = Windows8Palette.Palette.AccentColor;
                //add lines to the grid - windows 8 theme is a bit to white!
                if (StyleManager.ApplicationTheme.ToString() == "Windows8")
                {
                    //if we had any grids then add the lines like so
                    //radGridView1.VerticalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    //radGridView1.HorizontalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    //radGridView1.GridLinesVisibility = Telerik.Windows.Controls.GridView.GridLinesVisibility.Both;
                }
            }

            // hide the data tab
            tabData.Visibility = System.Windows.Visibility.Hidden;


            _d = Globals.ThisAddIn.getData();

            _templateid = "";
            _matterid = "";
            _versionid = "";
            _attachmentid = "";
            _attachedmode = true;
            _versionclonenewdocpath = "";

            _clauses = new Dictionary<string, FrameworkElement>();
            _elements = new Dictionary<string, FrameworkElement>();
        }



        public void LoadDataTab(string FileName, string ParentType, string ParentId)
        {

            // hide the data tab
            tabData.Visibility = System.Windows.Visibility.Visible;

            _parentType = ParentType;
            _parentId = ParentId;

            tbDocumentName.Text = System.IO.Path.GetFileName(FileName);
            tbDocumentName.IsReadOnly = true;

            btnSave.IsEnabled = false;

            _filename = System.IO.Path.GetFileName(FileName);

            BuildSidebar();
            this.SizeChanged += new SizeChangedEventHandler(Fields_SizeChanged);
        }

        public SForceEditSideBar2(string AttachmentId, string FileName, string ParentType, string ParentId)
        {


            _d = Globals.ThisAddIn.getData();

            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            if (StyleManager.ApplicationTheme.ToString() == "Windows8" || StyleManager.ApplicationTheme.ToString() == "Expression_Dark")
            {
                _setgbborder = true;
                _gbborder = Windows8Palette.Palette.AccentColor;
                //add lines to the grid - windows 8 theme is a bit to white!
                if (StyleManager.ApplicationTheme.ToString() == "Windows8")
                {
                    //if we had any grids then add the lines like so
                    //radGridView1.VerticalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    //radGridView1.HorizontalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    //radGridView1.GridLinesVisibility = Telerik.Windows.Controls.GridView.GridLinesVisibility.Both;
                }
            }


            // No Clause details
            // hide the clause tab and select the data tab
            this.tabClause.Visibility = System.Windows.Visibility.Collapsed;
            this.tabData.IsSelected = true;

            // show the export button on the data tab
            this.Save.Visibility = System.Windows.Visibility.Visible;

            _attachmentid = AttachmentId;

            tbDocumentName.Text = System.IO.Path.GetFileName(FileName);
            tbDocumentName.IsReadOnly = true;

            btnSave.IsEnabled = false;

            _filename = System.IO.Path.GetFileName(FileName);
            _parentType = ParentType;
            _parentId = ParentId;



            BuildSidebar();
            this.SizeChanged += new SizeChangedEventHandler(Fields_SizeChanged);

            Globals.ThisAddIn.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);

            /* --- clause choice stuff ---*/
            _templateid = "";
            _matterid = "";
            _versionid = "";

            _clauses = new Dictionary<string, FrameworkElement>();
            _elements = new Dictionary<string, FrameworkElement>();

        }

        //Clean up the Save


        ~SForceEditSideBar2()
        {
            Globals.ThisAddIn.Application.DocumentBeforeSave -= new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        private void BuildSidebar()
        {

            //Create an Expander for Document/Matter/Request - do them in order so we know the Id of the next one ...
            Fields.Children.Clear();



            if (this._d.demoinstance == "general" || this._d.demoinstance == "isda")
            {
                // assume that the parent is the Document
                _sDocumentObjectDef = new SForceEdit.SObjectDef("Document__c");
                if (_sDocumentObjectDef != null)
                {
                    GenerateFields(_sDocumentObjectDef);
                    LoadData(_parentId, _sDocumentObjectDef);
                }
            }
            else
            {
                // hack for the Document version if the parent is document OR version doesn't exit
                _sDocumentObjectDef = new SForceEdit.SObjectDef("Version__c");

                if (_parentType == "Document__c" || _sDocumentObjectDef == null)
                {
                    _sDocumentObjectDef = new SForceEdit.SObjectDef("Document__c");
                    GenerateFields(_sDocumentObjectDef);
                    LoadData(_parentId, _sDocumentObjectDef);
                }
                else
                {

                    if (_parentType == "Version__c")
                    {
                        GenerateFields(_sDocumentObjectDef);
                        LoadData(_parentId, _sDocumentObjectDef);

                        //If we have a Matter add that in
                        if (_DocumentRow != null && _DocumentRow.Table.Columns.Contains("Matter__c"))
                        {
                            string id = _DocumentRow["Matter__c"].ToString();
                            _sMatterObjectDef = new SForceEdit.SObjectDef("Matter__c");
                            GenerateFields(_sMatterObjectDef);
                            LoadData(id, _sMatterObjectDef);

                            //Add in the Activities
                            // _sActivityObjectDef = new SForceEdit.SObjectDef("Task");
                            // AddGrid(_sActivityObjectDef);
                            // GenerateFields(_sActivityObjectDef);
                            // LoadData(id, _sActivityObjectDef);

                        }

                        //If we have a Request add that in
                        if (_MatterRow != null && _MatterRow.Table.Columns.Contains("Request__c"))
                        {
                            string id = _MatterRow["Request__c"].ToString();
                            _sRequestObjectDef = new SForceEdit.SObjectDef("Request__c");
                            GenerateFields(_sRequestObjectDef);
                            LoadData(id, _sRequestObjectDef);
                        }
                    }
                    else if (_parentType == "Matter__c")
                    {
                        _sMatterObjectDef = new SForceEdit.SObjectDef("Matter__c");
                        GenerateFields(_sMatterObjectDef);
                        LoadData(_parentId, _sMatterObjectDef);

                        //If we have a Request add that in
                        if (_MatterRow != null && _MatterRow.Table.Columns.Contains("Request__c"))
                        {
                            string id = _MatterRow["Request__c"].ToString();
                            _sRequestObjectDef = new SForceEdit.SObjectDef("Request__c");
                            GenerateFields(_sRequestObjectDef);
                            LoadData(id, _sRequestObjectDef);
                        }
                    }
                    else if (_parentType == "Request__c")
                    {
                        _sRequestObjectDef = new SForceEdit.SObjectDef("Request__c");
                        GenerateFields(_sRequestObjectDef);
                        LoadData(_parentId, _sRequestObjectDef);
                    }
                }
            }
        }

        private void GenerateFields(SForceEdit.SObjectDef sObj)
        {

            sfPartner.DescribeSObjectResult dsr = _d.GetSObject(sObj.Name);
            sObj.Label = dsr.label;
            sObj.PluralLabel = dsr.labelPlural;
            if (_setgbborder) sObj.SetGBBorder(_gbborder);

            sObj.BuildCompactLayouts(_d, FieldChanged, SalesforcePressed, OpenPressed);

        }

        private void AddGrid(SForceEdit.SObjectDef sObj)
        {

            //Create the Grid
            RadGridView radGridView1 = new RadGridView();
            radGridView1.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            radGridView1.IsFilteringAllowed = false;
            radGridView1.IsReadOnly = true;
            radGridView1.ShowGroupPanel = false;

            if (_setgbborder) sObj.SetGBBorder(_gbborder);
            if (StyleManager.ApplicationTheme.ToString() == "Windows8")
            {
                radGridView1.VerticalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                radGridView1.HorizontalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                radGridView1.GridLinesVisibility = Telerik.Windows.Controls.GridView.GridLinesVisibility.Both;
            }

            //TODO need to bind the events as well

            radGridView1.Columns.Clear();
            radGridView1.AutoGenerateColumns = false;
            sObj.AddColumns(radGridView1);

            StackPanel spAct1 = new StackPanel();
            spAct1.Tag = sObj.Name;
            spAct1.Children.Add(radGridView1);
            Fields.Children.Add(spAct1);

        }

        private void LoadData(string Id, SForceEdit.SObjectDef sObj)
        {
            sObj.Id = Id;
            DataReturn dr = _d.GetData(sObj);
            AxiomIRISRibbon.Utility.HandleData(dr);
            DataRow r = null;
            if (dr.dt.Rows.Count > 0)
            {
                r = dr.dt.Rows[0];
            }


            //WE ONLY LOAD the data once with the sidebar so just add the panels here to 
            //the main stack panel and then update the data
            if (r != null && r.Table.Columns.IndexOf("RecordTypeId") >= 0)
            {
                string rid = r["RecordTypeId"].ToString();
                sfPartner.RecordTypeMapping m = sObj.RecordTypeMapping[rid];
                StackPanel sp = sObj.SideBarLayouts[m.layoutId];
                Fields.Children.Add(sp);
                //FieldContent.Content = sp;
                UpdateTextWidth();
            }
            else
            {
                string rid = sObj.DefaultRecordType;
                // if the layout doesn't contain the default then just use the first one - this shouldn't happen
                if (!sObj.RecordTypeMapping.ContainsKey(sObj.DefaultRecordType)) rid = sObj.RecordTypeMapping.ElementAt(0).Key;
                sfPartner.RecordTypeMapping m = sObj.RecordTypeMapping[rid];
                StackPanel sp = sObj.SideBarLayouts[m.layoutId];
                Fields.Children.Add(sp);
                //FieldContent.Content = sp;
                UpdateTextWidth();
            }

            if (sObj.Name == "Version__c")
            {
                _DocumentRow = r;
                SForceEdit.Utility.UpdateForm(FindStackPanel("Version__c"), _DocumentRow);
            }
            else if (sObj.Name == "Matter__c")
            {
                _MatterRow = r;
                SForceEdit.Utility.UpdateForm(FindStackPanel("Matter__c"), _MatterRow);
            }
            else if (sObj.Name == "Request__c")
            {
                _RequestRow = r;
                SForceEdit.Utility.UpdateForm(FindStackPanel("Request__c"), _RequestRow);
            }
            else if (sObj.Name == "Document__c")
            {
                _RequestRow = r;
                SForceEdit.Utility.UpdateForm(FindStackPanel("Document__c"), _RequestRow);
            }

            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }

        void Fields_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateTextWidth();
        }

        void UpdateTextWidth()
        {
            //Resize all the children
            for (int i = 0; i < Fields.Children.Count; i++)
            {
                //Fields then StackPanel for the Object then the Expander
                StackPanel s1 = (StackPanel)Fields.Children[i];
                if (s1.Children.Count > 0 && s1.Children[0].GetType() == typeof(Telerik.Windows.Controls.RadExpander))
                {
                    Telerik.Windows.Controls.RadExpander gb = (Telerik.Windows.Controls.RadExpander)s1.Children[0];
                    Grid g = (Grid)gb.Content;

                    int cols = g.ColumnDefinitions.Count;
                    double width = (FieldContent.ActualWidth / (cols / 2)) - 140;
                    if (width < 80) width = 80;

                    for (int j = 0; j < g.Children.Count; j++)
                    {

                        if (g.Children[j].GetType() == typeof(TextBox)) { ((TextBox)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(RadComboBox)) { ((RadComboBox)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(RadDatePicker)) { ((RadDatePicker)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(RadDateTimePicker)) { ((RadDateTimePicker)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(RadNumericUpDown)) { ((RadNumericUpDown)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(RadAutoCompleteBox)) { ((RadAutoCompleteBox)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(ScrollViewer)) { ((ScrollViewer)g.Children[j]).Width = width; }
                        else if (g.Children[j].GetType() == typeof(SForceEdit.AxSearchBox))
                        {
                            ((SForceEdit.AxSearchBox)g.Children[j]).Width = width;

                        }
                    }
                }
            }
        }

        
        void Application_DocumentBeforeSave(Word.Document doc, ref bool SaveAsUI, ref bool Cancel)
        {
            if (Globals.ThisAddIn.GetDocId(Globals.ThisAddIn.Application.ActiveDocument) == _attachmentid)
            {
                //Remove the property before saving and reapply - could jsut leave it
                //Globals.ThisAddIn.DeleteDocId(Globals.ThisAddIn.Application.ActiveDocument);
                SaveDocument();
                //Globals.ThisAddIn.AddDocId(Globals.ThisAddIn.Application.ActiveDocument, "attachmentid", _attachmentid);
                SaveAsUI = false;
                Cancel = true;
            }
        }

        private void SaveDocument()
        {
            //Save to salesforce
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string filenamenoext = System.IO.Path.GetFileNameWithoutExtension(_filename);

            //save this to a scratch file
            Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
            string filename = AxiomIRISRibbon.Utility.SaveTempFile(filenamenoext);
            doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatXMLDocument);

            //Save a copy!
            Globals.ThisAddIn.ProcessingUpdate("Save Copy");
            string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(filenamenoext + "X");
            Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: false);
            dcopy.SaveAs2(filenamecopy, Word.WdSaveFormat.wdFormatXMLDocument);

            var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
            docclose.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

            //Now 
            Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
            DataReturn dr = _d.UpdateFile(_attachmentid,"", filenamecopy);

            Globals.ThisAddIn.ProcessingStop("Stop");
        }

        private StackPanel FindStackPanel(string sObjName)
        {
            StackPanel ret = null;
            for (int i = 0; i < Fields.Children.Count; i++)
            {
                StackPanel s1 = (StackPanel)Fields.Children[i];
                if (s1.Tag != null && s1.Tag.ToString() == sObjName)
                {
                    return s1;
                }
            }
            return ret;
        }



        void FieldChanged()
        {
            StackPanel flds;
            bool changes = false;
            //Document
            if (_DocumentRow != null)
            {
                flds = FindStackPanel("Version__c");
                _DocumentRow.BeginEdit();
                if (SForceEdit.Utility.UpdateRow(flds, _DocumentRow)) changes = true;
                _DocumentRow.CancelEdit();
            }

            //Matter
            if (_MatterRow != null)
            {
                flds = FindStackPanel("Matter__c");
                _MatterRow.BeginEdit();
                if (SForceEdit.Utility.UpdateRow(flds, _MatterRow)) changes = true;
                _MatterRow.CancelEdit();
            }

            //Request
            if (_RequestRow != null)
            {
                flds = FindStackPanel("Request__c");
                _RequestRow.BeginEdit();
                if (SForceEdit.Utility.UpdateRow(flds, _RequestRow)) changes = true;
                _RequestRow.CancelEdit();
            }


            if (changes)
            {
                btnSave.IsEnabled = true;
                btnCancel.IsEnabled = true;
            }
            else
            {
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
            }
        }

        void SalesforcePressed(string sObjectType)
        {

            Uri temp = null;
            if (sObjectType == "Version__c" && _DocumentRow != null)
            {
                temp = new Uri(_sDocumentObjectDef.Url.Replace("{ID}", _DocumentRow["Id"].ToString()));
            }
            if (sObjectType == "Matter__c" && _MatterRow != null)
            {
                temp = new Uri(_sMatterObjectDef.Url.Replace("{ID}", _MatterRow["Id"].ToString()));
            }
            if (sObjectType == "Request__c" && _RequestRow != null)
            {
                temp = new Uri(_sRequestObjectDef.Url.Replace("{ID}", _RequestRow["Id"].ToString()));
            }


            if (temp != null)
            {

                string rooturl = temp.Scheme + "://" + temp.Host;
                string frontdoor = rooturl + "/secur/frontdoor.jsp?sid=" + _d.GetSessionId();
                string redirect = frontdoor + "&retURL=" + temp.PathAndQuery;
                System.Diagnostics.Process.Start(redirect);
            }


        }

        void OpenPressed(string sObjectType)
        {
            if (sObjectType == "Version__c" && _DocumentRow != null)
            {
                Globals.ThisAddIn.OpenZoomEditWindow(sObjectType, _DocumentRow["Id"].ToString());
            }
            if (sObjectType == "Matter__c" && _MatterRow != null)
            {
                Globals.ThisAddIn.OpenZoomEditWindow(sObjectType, _MatterRow["Id"].ToString());
            }
            if (sObjectType == "Request__c" && _RequestRow != null)
            {
                Globals.ThisAddIn.OpenZoomEditWindow(sObjectType, _RequestRow["Id"].ToString());
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Saving ...";

            _DocumentChanges = SForceEdit.Utility.UpdateRow(FindStackPanel("Version__c"), _DocumentRow);
            _MatterChanges = SForceEdit.Utility.UpdateRow(FindStackPanel("Matter__c"), _MatterRow);
            _RequestChanges = SForceEdit.Utility.UpdateRow(FindStackPanel("Request__c"), _RequestRow);

            if (_DocumentChanges || _MatterChanges || _RequestChanges)
            {
                _saveBackgroundWorker = new BackgroundWorker();
                _saveBackgroundWorker.DoWork += (obj, ev) => saveWorkerDoWork(obj, ev);
                _saveBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(saveBackgroundWorker_RunWorkerCompleted);
                _saveBackgroundWorker.RunWorkerAsync();
            }
            else
            {
                //shouldn't really happen
                bsyInd.IsBusy = false;
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
            }

        }

        void saveBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _saveBackgroundWorker.DoWork -= (obj, ev) => saveWorkerDoWork(obj, ev);
            _saveBackgroundWorker.RunWorkerCompleted -= saveBackgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;

            DataReturn dr = (DataReturn)e.Result;
            AxiomIRISRibbon.Utility.HandleData(dr);

            if (dr.success)
            {
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
            }


        }


        void saveWorkerDoWork(object sender, DoWorkEventArgs e)
        {

            DataReturn dr = new DataReturn();

            if (_DocumentChanges)
            {
                DataReturn tempdr = _d.Save(_sDocumentObjectDef, _DocumentRow);
                if (!tempdr.success)
                {
                    _versionid = dr.id;

                    dr.success = false;
                    dr.errormessage += tempdr.errormessage;
                }

            }

            if (_MatterChanges)
            {
                DataReturn tempdr = _d.Save(_sMatterObjectDef, _MatterRow);
                if (!tempdr.success)
                {
                    _MatterRow.CancelEdit();
                    dr.success = false;
                    dr.errormessage += tempdr.errormessage;
                }
                else
                {
                    _MatterRow.EndEdit();
                }
            }

            if (_RequestChanges)
            {
                DataReturn tempdr = _d.Save(_sRequestObjectDef, _RequestRow);
                if (!tempdr.success)
                {
                    _RequestRow.CancelEdit();
                    dr.success = false;
                    dr.errormessage += tempdr.errormessage;
                }
                else
                {
                    _RequestRow.EndEdit();
                }
            }


            e.Result = dr;
        }


        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {

            // Revert
            SForceEdit.Utility.UpdateForm(FindStackPanel("Version__c"), _DocumentRow);
            SForceEdit.Utility.UpdateForm(FindStackPanel("Matter__c"), _MatterRow);
            SForceEdit.Utility.UpdateForm(FindStackPanel("Request__c"), _RequestRow);


            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }


        // Save as a regular file!
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string filenamenoext = System.IO.Path.GetFileNameWithoutExtension(_filename);

            //save this to a scratch file so we can copy
            Globals.ThisAddIn.DeleteDocId(Globals.ThisAddIn.Application.ActiveDocument);
            Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
            string filename = AxiomIRISRibbon.Utility.SaveTempFile(filenamenoext);
            doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatXMLDocument);

            //Create a copy!
            Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: true);
            dcopy.Activate();
            try
            {
                dcopy.Save();
            }
            catch (Exception)
            {
            }

        }



        /* --- clause choice stuff ---*/

        public class PlaybookTag
        {
            private string _id;
            private string _html;
            private string _type;

            public string id
            {
                get { return _id; }
                set { _id = value; }
            }
            public string html
            {
                get { return _html; }
                set { _html = value; }
            }
            public string type
            {
                get { return _type; }
                set { _type = value; }
            }
        }


        private class Element
        {
            private string _docelementid;
            private string _conceptid;

            //Template values for the element and the clauseelement link
            private string _templateelementid;
            private string _templateclauseelementid;
            private string _templateelementname;

            //Option string 
            private string[] _options;
            //Option values for checkbox
            private string _option1;
            private string _option2;

            private string _type;
            private string _controltype;
            private string _format;

            private string _default;
           
            //Orignal Value
            private string _originalvalue;

            //Last selected for dropdown
            private string _lastselected;

            public string docelementid
            {
                get { return _docelementid; }
                set { _docelementid = value; }
            }
            public string templateelementid
            {
                get { return _templateelementid; }
                set { _templateelementid = value; }
            }
            public string templateclauseelementid
            {
                get { return _templateclauseelementid; }
                set { _templateclauseelementid = value; }
            }
            public string templateelementname
            {
                get { return _templateelementname; }
                set { _templateelementname = value; }
            }
            public string conceptid
            {
                get { return _conceptid; }
                set { _conceptid = value; }
            }
            public string[] options
            {
                get { return _options; }
                set { _options = value; }
            }
            public string option1
            {
                get { return _option1; }
                set { _option1 = value; }
            }
            public string option2
            {
                get { return _option2; }
                set { _option2 = value; }
            }

            public string format
            {
                get { return _format; }
                set { _format = value; }
            }
            public string type
            {
                get { return _type; }
                set { _type = value; }
            }
            public string controltype
            {
                get { return _controltype; }
                set { _controltype = value; }
            }
            public string lastselected
            {
                get { return _lastselected; }
                set { _lastselected = value; }
            }
            public string originalvalue
            {
                get { return _originalvalue; }
                set { _originalvalue = value; }
            }

            public string defaultvalue
            {
                get { return _default; }
                set { _default = value; }
            }
        }


        private class ClauseRadio : RadioButton
        {

            //Template Fields
            private string _id;
            private string _name;
            private string _conceptid;
            private string _conceptname;
            private string _priority;
            private string _risk;
            private string _xml;
            private string _lastmodified;
            private int _number;
            private DataRow _dr;
            private string _approver;
            private bool _unlock;
            private string _unlockapprover;

            //Contract Fields            
            private string _documentclauseid;

            public ClauseRadio(string id, string name, string conceptid, string conceptname, int number, DataRow dr, string xml, string lastmodified, string approver, bool unlock, string unlockapprover, string description)
            {
                _id = id;
                _name = name;
                _conceptid = conceptid;
                _conceptname = conceptname;
                _priority = priority;
                _risk = risk;
                _xml = xml;
                _lastmodified = lastmodified;
                _number = number;
                _dr = dr;
                _approver = approver;
                _unlock = unlock;
                _unlockapprover = unlockapprover;

                Content = name;
                ToolTip = description;
                Margin = new Thickness(30, 5, 0, 5);
            }
            public string id
            {
                get { return _id; }
            }
            public string name
            {
                get { return _name; }
            }
            public string conceptid
            {
                get { return _conceptid; }
            }
            public string conceptname
            {
                get { return _conceptname; }
            }
            public int number
            {
                get { return _number; }
            }
            public string priority
            {
                get { return _priority; }
            }

            public string risk
            {
                get { return _risk; }
            }

            public DataRow dr
            {
                get { return _dr; }
            }

            public string xml
            {
                get { return _xml; }
                set { _xml = value; }
            }

            public string lastmodified
            {
                get { return _lastmodified; }
                set { _lastmodified = value; }
            }

            public string documentclauseid
            {
                get { return _documentclauseid; }
                set { _documentclauseid = value; }
            }

            public string approver
            {
                get { return _approver; }
                set { _approver = value; }
            }

            public bool unlock
            {
                get { return _unlock; }
                set { _unlock = value; }
            }

            public string unlockapprover
            {
                get { return _unlockapprover; }
                set { _unlockapprover = value; }
            }
        }


        public void BuildSideBarNewVersion(string TemplateId, string TemplateName, string TemplatePlaybookLink,string MatterId, string MatterName)
        {
            _matterid = MatterId;
            _templateid = TemplateId;
            _versionid = "";
            _attachmentid = "";

            // Get the new Version Name and Number
            string VersionName = "";
            string VersionNumber = "";
            DataReturn versionmax = _d.GetVersionMax(_matterid);
            string vmax = versionmax.dt.Rows[0][0].ToString();
            double vmaxint = 1;
            if (vmax != null)
            {
                try
                {
                    vmaxint = Convert.ToDouble(vmax) + 1;
                }
                catch (Exception)
                {

                }
            }
            VersionName = "Version " + vmaxint.ToString();
            VersionNumber = vmaxint.ToString();

            this.tbMatterName.Text = MatterName;
            this.tbVersionName.Text = VersionName;
            this.tbVersionNumber.Text = VersionNumber;

            // create the version - no cloning so very basic data
            DataReturn dr = Utility.HandleData(_d.SaveVersion(_versionid, _matterid, _templateid, VersionName, VersionNumber));
            if (!dr.success) return;
            _versionid = dr.id;

            this.LoadCompareMenu();
            this.BuildSideBar(TemplateId, TemplateName, TemplatePlaybookLink);            
            this.LoadDataTab(_d.contractfilename, "Version__c", _versionid);

            // once we have loaded the data tab populate the default element values - some of which may be read from the 
            // the data tab
            Globals.ThisAddIn.ProcessingUpdate("Set Default Clauses");
            this.SetDefaultClauses();
            Globals.ThisAddIn.ProcessingUpdate("Load Elements");
            this.LoadElementsFromDefault();
            Globals.ThisAddIn.ProcessingUpdate("Initiate Elements");
            this.InitiateElements();

        }


        // when we don't have a Matter - this is for testing the template 
        public void BuildSideNoVersion(string TemplateId, string TemplateName, string TemplatePlaybookLink)
        {
            _matterid = "";
            _templateid = TemplateId;
            _versionid = "";
            _attachmentid = "";

            this.tbMatterName.Text = "";
            this.tbVersionName.Text = "";
            this.tbVersionNumber.Text = "";

            this.BuildSideBar(TemplateId, TemplateName, TemplatePlaybookLink);
            
            // once we have loaded the data tab populate the default element values and default clauses
            // no DataTab so will show the first one or a blank value if they are formulas
            Globals.ThisAddIn.ProcessingUpdate("Set Default Clauses");
            if (!Globals.ThisAddIn.getDebug()) this.SetDefaultClauses();
            Globals.ThisAddIn.ProcessingUpdate("Load Elements");
            this.LoadElementsFromDefault();
            Globals.ThisAddIn.ProcessingUpdate("Initiate Elements");
            this.InitiateElements();

            // Switch off the menus that relly on having a version
            this.Compare.IsEnabled = false;
            this.NewVersion.IsEnabled = false;

        }

        public void BuildSideBarFromVersion(string VersionId,string AttachedMode,string AttachmentId)
        {
            this.SetAttachedMode(AttachedMode);
            _attachmentid = AttachmentId;
            
            // get the required data from the version
            DataReturn dr = Utility.HandleData(_d.GetVersion(VersionId));
            if (!dr.success) return;

            if (dr.dt.Rows.Count != 1)
            {
                MessageBox.Show("Cannot find the version");
                return;
            }

            _versionid = VersionId;
            string VersionName = dr.dt.Rows[0]["Name"].ToString();

            if (dr.dt.Rows[0]["Template__c"] != null && dr.dt.Rows[0]["Template__c"].ToString() != "")
            {

                string TemplateId = dr.dt.Rows[0]["Template__c"].ToString();
                string TemplateName = dr.dt.Rows[0]["Template__r_Name"].ToString();
                string TemplatePlaybookLink = dr.dt.Rows[0]["Template__r_PlaybookLink__c"].ToString();
                this.tbTemplateName.Text = dr.dt.Rows[0]["Template__r_Name"].ToString();
                this.tbVersionName.Text = dr.dt.Rows[0]["Name"].ToString();
                this.tbVersionNumber.Text = dr.dt.Rows[0]["Version_Number__c"].ToString();

                // hack to get it to work with the document version
                if (dr.dt.Columns.Contains("Matter__r_Name"))
                {
                    _matterid = dr.dt.Rows[0]["Matter__c"].ToString();
                    this.tbMatterName.Text = dr.dt.Rows[0]["Matter__r_Name"].ToString();
                }
                else
                {
                    if (dr.dt.Columns.Contains("Request2__c"))
                    {
                        // general demo
                        _matterid = dr.dt.Rows[0]["Request2__c"].ToString();
                        this.lbMatter.Content = "Request";
                        this.lbVersion.Content = "Document";
                        this.tbMatterName.Text = dr.dt.Rows[0]["Request2__r_Name"].ToString();
                    }
                    else
                    {
                        // isda demo
                        _matterid = dr.dt.Rows[0]["Version__c"].ToString();
                        this.lbMatter.Content = "Version";
                        this.lbVersion.Content = "Document";
                        this.tbMatterName.Text = dr.dt.Rows[0]["Version__r_Name"].ToString();
                    }
                }
                this.LoadCompareMenu();
                this.BuildSideBar(TemplateId, TemplateName, TemplatePlaybookLink);
            } else {
                // we have no template - template could have been deleted
                MessageBox.Show("The Template that this contract was based on has been removed.");
                this._doc = Globals.ThisAddIn.Application.ActiveDocument;
                Globals.ThisAddIn.RemoveContentControls(this._doc);

                // No Clause details
                // hide the clause tab and select the data tab
                this.tabClause.Visibility = System.Windows.Visibility.Collapsed;
                this.tabData.IsSelected = true;

                // show the export button on the data tab
                this.Save.Visibility = System.Windows.Visibility.Visible;
                _attachmentid = AttachmentId;
                btnSave.IsEnabled = false;


            }

            Globals.ThisAddIn.ProcessingUpdate("Load Contract Data");
            this.LoadContractData(VersionId, VersionName);
        }


        public void LoadCompareMenu()
        {
            this.CompareContent.Items.Clear();

            // Load all the other version records
            DataReturn dr = Utility.HandleData(_d.GetVersionFromMatter(_matterid));
            if (!dr.success) return;


            foreach (DataRow r in dr.dt.Rows)
            {
                if (r["Id"].ToString() != _versionid)
                {
                    RadMenuItem mi = new RadMenuItem();

                    string versionnum = "";
                    if (r["Version_Number__c"].ToString() != "")
                    {
                        try
                        {
                            versionnum = Convert.ToDecimal(r["Version_Number__c"]).ToString("0") + "-";
                        }
                        catch (Exception)
                        {

                        }
                    }
                    
                    mi.Header = versionnum + r["Name"].ToString();

                    mi.Tag = r["Id"].ToString();
                    this.CompareContent.Items.Add(mi);
                }

                // if none then hide
                if (this.CompareContent.Items.Count == 0)
                {
                    this.Compare.Visibility = System.Windows.Visibility.Collapsed;
                }
                else
                {
                    this.Compare.Visibility = System.Windows.Visibility.Visible;
                }
            }
        }



        public void BuildSideBar(string TemplateId, string TemplateName, string TemplatePlaybookLink)
        {

            Globals.ThisAddIn.ProcessingStart("Build Side Bar");

            try
            {
                Data d = Globals.ThisAddIn.getData();

                _templateid = TemplateId;
                _templateplaybooklink = TemplatePlaybookLink;
                tbTemplateName.Text = TemplateName;

                if (_templateplaybooklink != "")
                {
                    ToolTip tt = new System.Windows.Controls.ToolTip();
                    tt.Content = "Open the Template Playbook link: " + TemplatePlaybookLink;
                    btnTemplatePlaybook.ToolTip = tt;
                }
                else
                {
                    // mark with no link
                    btnTemplatePlaybook.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                }


                // if this is an attached doc then to get the order we step through the document
                // and get the controls in order
                // if this is an unattached document then we get the template and get the order from that

                _doc = Globals.ThisAddIn.Application.ActiveDocument;

                Word.Document orderdoc = null;
                Word.Range orderrng = null;

                if (_attachedmode)
                {
                    orderdoc = Globals.ThisAddIn.Application.ActiveDocument;
                    object start = orderdoc.Content.Start;
                    object end = orderdoc.Content.End;
                    orderrng = orderdoc.Range(ref start, ref end);
                }
                else
                {
                    string filename = Utility.SaveTempFile(TemplateId);
                    Globals.ThisAddIn.ProcessingUpdate("Download Template File From SForce");
                    DataReturn orderdr = Utility.HandleData(d.GetTemplateFile(TemplateId, filename));
                    if (!orderdr.success) return;
                    filename = orderdr.strRtn;
                    orderdoc = Globals.ThisAddIn.Application.Documents.Open(filename, Visible: false);
                    object start = orderdoc.Content.Start;
                    object end = orderdoc.Content.End;
                    orderrng = orderdoc.Range(ref start, ref end);
                }

                _clauses = new Dictionary<string, FrameworkElement>();
                _elements = new Dictionary<string, FrameworkElement>();
                //Clear any old ones
                Questions.Children.Clear();

                lbApprovals.Visibility = System.Windows.Visibility.Hidden;
                btnApprovals.Visibility = System.Windows.Visibility.Hidden;
                this.rdTopPanel.Height = new GridLength(85);
                Globals.Ribbons.Ribbon1.Approval(false);

                //Make the default the first one for now
                int num = 0;

                //Get all the clauses and get all the elements at the start
                DataReturn dr = Utility.HandleData(_d.GetTemplateClauses(TemplateId, ""));
                if (!dr.success) return;
                DataTable allclauses = dr.dt;

                //Generate a list of the ClauseIds so we can get all the elemets at one
                //short term solution to cut down on the API calls     
                List<string> clauseids = new List<string>();
                foreach (DataRow r in allclauses.Rows)
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

                dr = Utility.HandleData(_d.GetMultipleClauseElements(clausefilter));
                if (!dr.success) return;
                DataTable allelements = dr.dt;


                //Now step through all the Contact Controls and update the XML so we get the newest clauses
                Globals.ThisAddIn.ProcessingUpdate("Step Through Clauses");
                foreach (Word.ContentControl cc in orderrng.ContentControls)
                {
                    if (cc.Tag != null)
                    {
                        string tag = cc.Tag;
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept")
                        {

                            DataRow[] clauses = allclauses.Select("Clause__r_Concept__r_Id='" + Convert.ToString(taga[1]) + "'", "Order__c,Clause__r_Name");

                            string conceptid = "";
                            string conceptname = "";

                            string pbInfo = "";
                            string pbClient = "";

                            if (clauses.Length > 0)
                            {
                                conceptid = Convert.ToString(clauses[0]["Clause__r_Concept__r_Id"]);
                                conceptname = clauses[0]["Clause__r_Concept__r_Name"].ToString();
                                pbInfo = clauses[0]["Clause__r_Concept__r_PlayBookInfo__c"].ToString();
                                pbClient = clauses[0]["Clause__r_Concept__r_PlayBookClient__c"].ToString();
                            }

                            //Add in the Concept Expander


                            Grid gExp = new Grid();
                            Expander newExp = new Expander();
                            newExp.Header = conceptname;
                            newExp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                            newExp.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xEC, 0xED, 0xED));
                            newExp.Name = "exp" + conceptid;
                            newExp.Tag = conceptid;
                            newExp.IsExpanded = true;

                            int paddingforlockbutton = 30;
                            if (!_attachedmode)
                            {
                                paddingforlockbutton = 6;
                            }

                            Button lExp1 = new Button();
                            Style style = this.FindResource("LinkButton") as Style;
                            lExp1.Style = style;
                            if (pbInfo == "") lExp1.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                            lExp1.Margin = new Thickness(0, 8, paddingforlockbutton, 0);
                            lExp1.ToolTip = ConvertHTMLToToolTip(pbInfo);
                            lExp1.Content = "Info";
                            lExp1.Height = 28;

                            PlaybookTag pb = new PlaybookTag();
                            pb.id = conceptid;
                            pb.html = pbInfo;
                            pb.type = "Info";

                            lExp1.Tag = pb;
                            lExp1.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                            lExp1.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            lExp1.Click += new RoutedEventHandler(lExp1_Click);

                            Button lExp2 = new Button();
                            lExp2.Style = style;
                            if (pbClient == "") lExp2.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                            lExp2.Margin = new Thickness(0, 8, paddingforlockbutton + 30, 0);
                            lExp2.ToolTip = ConvertHTMLToToolTip(pbClient);
                            lExp2.Content = "Client";

                            pb = new PlaybookTag();
                            pb.id = conceptid;
                            pb.html = pbClient;
                            pb.type = "Client";

                            lExp2.Tag = pb;
                            lExp2.Height = 28;
                            lExp2.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                            lExp2.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            lExp2.Click += new RoutedEventHandler(lExp2_Click);

                            Button unlock = new Button();
                            unlock.Margin = new Thickness(0, 2, 2, 0);
                            unlock.ToolTip = "Unlock Clause";

                            Image icon = new Image();
                            icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                            unlock.Content = icon;
                            unlock.Name = "unlock" + conceptid;
                            unlock.Tag = conceptid;
                            unlock.Height = 22;
                            unlock.Width = 22;
                            unlock.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                            unlock.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            unlock.Click += new RoutedEventHandler(unlock_Click);

                            gExp.Children.Add(newExp);
                            gExp.Children.Add(lExp1);
                            gExp.Children.Add(lExp2);
                            gExp.Children.Add(unlock);

                            // if we are unattached then hide the lock cause it doesn't make that much sense
                            if (!_attachedmode) unlock.Visibility = System.Windows.Visibility.Hidden;


                            StackPanel spCl = new StackPanel();


                            foreach (DataRow r in clauses)
                            {

                                StackPanel spEl = new StackPanel();
                                spEl.Tag = "elementsp";
                                spEl.Margin = new Thickness(35, 5, 5, 5);
                                StackPanel spRb = new StackPanel();
                                spRb.Tag = "rbsp";

                                //Get the XML for this clause                                
                                string xml = "";
                                string clauseid = Convert.ToString(r["Clause__r_Id"]);
                                string lastmodified = Convert.ToString(r["Clause__r_LastModifiedDate"]);
                                lastmodified = lastmodified.Substring(0, 16);

                                // ----- If the Clause is the one in the template/contract then get it from the control if timestamp matches
                                // ----- if the cause is unlocked then it will be set to "Unlocked" and we need to load
                                // ----- otherwise load the clause
                                if (_attachedmode)
                                {
                                    bool selectedclause = false;
                                    if (taga.Length > 3)
                                    {
                                        // clause matches and timestamp or if timestap is set to 0000 which indicates its been loaded
                                        if (taga[2] == clauseid && (lastmodified == taga[3] || taga[3] == "0000"))
                                        {
                                            selectedclause = true;
                                        }
                                    }

                                    if (selectedclause)
                                    {
                                        Globals.ThisAddIn.ProcessingUpdate("Take XML from Doc");
                                        // get the xml from the doc
                                        xml = Globals.ThisAddIn.GetContractClauseXML(_doc, conceptid);
                                    }

                                    if (xml == "")
                                    {

                                        Globals.ThisAddIn.ProcessingUpdate("Get the Clause Template File from SF for:" + r["Clause__r_Name"].ToString());
                                        string filename = Utility.SaveTempFile(clauseid);
                                        filename = Utility.HandleData(_d.GetClauseFile(clauseid, filename)).strRtn;
                                        if (filename == "")
                                        {
                                            xml = "Sorry can't find clause";
                                        }
                                        else
                                        {
                                            //This is the bit that causes the flash - have to open and close the file
                                            Globals.ThisAddIn.Application.ScreenUpdating = false;
                                            Word.Document doc1 = Globals.ThisAddIn.Application.Documents.Open(filename, Visible: false);
                                            xml = doc1.WordOpenXML;
                                            var docclose = (Microsoft.Office.Interop.Word._Document)doc1;
                                            docclose.Close();
                                            Globals.ThisAddIn.Application.ScreenUpdating = true;
                                        }
                                    }
                                }

                                //Approvals - work if we have an approver
                                string desc = r["Clause__r_Description__c"].ToString();
                                string approver = "";
                                if (desc.Contains("Approver:"))
                                {
                                    int i1 = desc.IndexOf("Approver:") + "Approver:".Length;
                                    int i2 = desc.IndexOf("\n", i1);
                                    if (i2 == -1) i2 = desc.Length;
                                    approver = desc.Substring(i1, i2 - i1);
                                }

                                string unlockapprover = "";
                                if (desc.Contains("ApproverFreeText:"))
                                {
                                    int i1 = desc.IndexOf("ApproverFreeText:") + "ApproverFreeText:".Length;
                                    int i2 = desc.IndexOf("\n", i1);
                                    if (i2 == -1) i2 = desc.Length;
                                    unlockapprover = desc.Substring(i1, i2 - i1);
                                }

                                if (unlockapprover == "") unlockapprover = approver;

                                //Approval over ---

                                //Add in the radio button header

                                ClauseRadio rb1 = new ClauseRadio(Convert.ToString(r["Clause__r_Id"]), r["Clause__r_Name"].ToString(), conceptid, conceptname, num, r, xml, lastmodified, approver, false, unlockapprover, desc);
                                rb1.GroupName = conceptname;
                                rb1.Checked += new RoutedEventHandler(rb1_Checked);
                                rb1.GotFocus += new RoutedEventHandler(rb1_GotFocus);
                                spRb.Margin = new Thickness(5, 5, 5, 5);

                                spRb.Children.Add(rb1);



                                //Approvals put in the button ---------------------------------------------------
                                Button ApprovalButton = new Button();
                                ApprovalButton.Margin = new Thickness(2, 2, 2, 0);
                                // ApprovalButton.ToolTip = ConvertHTMLToToolTip(pbInfo);
                                ApprovalButton.Content = "Get Approval";
                                ApprovalButton.Tag = conceptid + "|" + conceptname + "|" + approver;
                                ApprovalButton.Height = 22;
                                ApprovalButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                                ApprovalButton.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                ApprovalButton.Click += new RoutedEventHandler(ApprovalButton_Click);

                                System.Windows.Controls.Label lbl = new System.Windows.Controls.Label();
                                lbl.Content = "This Clause Requires Approval from: " + approver;
                                lbl.Margin = new Thickness(4, 4 + (num * 27), 10, 0);
                                lbl.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                //lbl.Width = 120;

                                Grid gApproval = new Grid();
                                gApproval.Height = 32;
                                gApproval.Children.Add(lbl);
                                gApproval.Children.Add(ApprovalButton);

                                spEl.Children.Add(gApproval);

                                if (approver == "")
                                {
                                    gApproval.Visibility = System.Windows.Visibility.Collapsed;
                                }
                                //Approvals ---------------------------------------------------


                                //Add in the elements for this clause
                                //dr = _d.GetElements(Convert.ToString(r["Clause__r_Id"]));
                                //if (!dr.success) return;
                                //DataTable elements = dr.dt;

                                DataRow[] elements = allelements.Select("Clause__r_Id='" + Convert.ToString(r["Clause__r_Id"]) + "'");

                                if (elements.Length > 0)
                                {
                                    Globals.ThisAddIn.ProcessingUpdate("Add In Elements");

                                    foreach (DataRow er in elements)
                                    {
                                        //Add in the element control
                                        lbl = new System.Windows.Controls.Label();

                                        string lblstr = er["Element__r_Label__c"].ToString();
                                        if (lblstr == "") lblstr = er["Element__r_Name"].ToString();


                                        /* OLD TEXT BOX ONLY code
                                        TextBox tb = new TextBox();
                                        tb.Height = 23;
                                        tb.Margin = new Thickness(90, 4 + (num * 27), 10, 0);
                                        tb.Name = "tb" + er["Element__r_Name"].ToString();
                                        tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;

                                        string dflt = er["Element__r_DefaultValue__c"].ToString();
                                        if (dflt != "")
                                        {

                                            if (dflt == "=Now")
                                            {
                                                DateTime now = DateTime.Now;
                                                if (er["Element__r_Format__c"].ToString() != "")
                                                {
                                                    tb.Text = now.ToString(er["Element__r_Format__c"].ToString());
                                                }
                                                else
                                                {
                                                    tb.Text = now.ToString("d MMMM yyyy");
                                                }

                                            }
                                            else
                                            {
                                                tb.Text = er["Element__r_DefaultValue__c"].ToString();
                                            }
                                        }
                                        Element e = new Element();
                                        e.docelementid = "";
                                        e.templateelementid = er["Element__r_Id"].ToString();
                                        e.conceptid = Convert.ToString(r["Clause__r_Id"]); //hold the clause so we know what to open if they click
                                        e.templateclauseelementid = Convert.ToString(er["Id"]); //hold this so we don't update the same field when
                                        e.templateelementname = er["Element__r_Name"].ToString();
                                        tb.Tag = e;

                                        tb.TextChanged += new TextChangedEventHandler(element_TextChanged);
                                        tb.GotFocus += new RoutedEventHandler(tb_GotFocus);
                                        Grid g2 = new Grid();
                                        //StackPanel sp2 = new StackPanel();
                                        //g2.Orientation = Orientation.Horizontal;
                                        g2.Height = 32;
                                        g2.Children.Add(lbl);
                                        g2.Children.Add(tb);

                                        spEl.Children.Add(g2);
                                         * */

                                        //------------------- PULL THIS OUT SO IT IS MORE MODULAR! make it easier to add diferent types
                                        lbl.Content = lblstr + ":";
                                        lbl.Margin = new Thickness(4, 4 + (num * 27), 10, 0);
                                        lbl.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                        lbl.Width = 120;

                                        Grid g2 = new Grid();
                                        g2.Height = 32;

                                        if (er["Element__r_Type__c"].ToString() == "Picklist")
                                        {
                                            ComboBox cb = new ComboBox();
                                            //cb.Width = 200;
                                            cb.Height = 23;

                                            cb.IsEditable = true;
                                            cb.IsTextSearchEnabled = true;

                                            cb.Name = "cb" + er["Element__r_Name"].ToString().Replace(" ", ""); ;
                                            cb.Margin = new Thickness(120, 4 + (num * 27), 10, 0);
                                            cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;

                                            string options = er["Element__r_Options__c"].ToString().Replace("\r", "");
                                            string[] entries = options.Split('\n');

                                            cb.ItemsSource = entries;
                                            cb.SelectedValuePath = "Content";
                                            //cb.Tag = er["ID"].ToString() + "|" + er["ClauseElementId"].ToString() + "|" + conceptid + "|";

                                            string dflt = er["Element__r_DefaultValue__c"].ToString();
                                            // if (dflt != "") cb.SelectedItem = dflt;

                                            Element e = new Element();
                                            e.docelementid = "";
                                            e.type = er["Element__r_Type__c"].ToString();
                                            e.controltype = "ComboBox";
                                            e.format = er["Element__r_Format__c"].ToString();
                                            e.templateelementid = er["Element__r_Id"].ToString();
                                            e.conceptid = conceptid; //hold the clause so we know what to open if they click
                                            e.templateclauseelementid = Convert.ToString(er["Id"]); //hold this so we don't update the same field when
                                            e.templateelementname = er["Element__r_Name"].ToString();
                                            e.options = entries;
                                            e.defaultvalue = dflt;
                                            cb.Tag = e;

                                            cb.GotFocus += new RoutedEventHandler(cb_GotFocus);
                                            cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                                            cb.LostFocus += new RoutedEventHandler(cb_LostFocus);

                                            _elements.Add(e.templateclauseelementid, cb);

                                            g2.Children.Add(lbl);
                                            g2.Children.Add(cb);


                                        }
                                        else if (er["Element__r_Type__c"].ToString() == "Checkbox")
                                        {
                                            CheckBox cbox = new CheckBox();
                                            //cb.Width = 200;
                                            cbox.Height = 23;

                                            cbox.Name = "cbox" + er["Element__r_Name"].ToString().Replace(" ", "");
                                            cbox.Margin = new Thickness(10, 4 + (num * 27), 10, 0);
                                            cbox.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                            cbox.Content = lblstr;

                                            string options = er["Element__r_Options__c"].ToString().Replace("\r", "");
                                            string[] entries = options.Split('\n');
                                            string opt1 = "", opt2 = "";
                                            if (entries.Length >= 1) opt1 = entries[0];
                                            if (entries.Length >= 2) opt2 = entries[1];

                                            string dflt = er["Element__r_DefaultValue__c"].ToString().ToLower();
                                            // if (dflt != "")
                                            // {
                                            //     if (dflt == "y" || dflt == "yes" || dflt == "true" || dflt == "t" || dflt == opt1.ToLower())
                                            //     {
                                            //         cbox.IsChecked = true;
                                            //     }
                                            // }



                                            //cbox.Tag = er["ID"].ToString() + "|" + er["ClauseElementId"].ToString() + "|" + conceptid + "|" + opt1 + "|" + opt2 ;                                    
                                            Element e = new Element();
                                            e.docelementid = "";
                                            e.templateelementid = er["Element__r_Id"].ToString();
                                            e.type = er["Element__r_Type__c"].ToString();
                                            e.controltype = "CheckBox";
                                            e.format = er["Element__r_Format__c"].ToString();
                                            e.conceptid = conceptid; //hold the clause so we know what to open if they click
                                            e.templateclauseelementid = Convert.ToString(er["Id"]); //hold this so we don't update the same field when
                                            e.templateelementname = er["Element__r_Name"].ToString();
                                            e.options = entries;
                                            e.option1 = opt1;
                                            e.option2 = opt2;
                                            e.defaultvalue = dflt;
                                            cbox.Tag = e;

                                            cbox.GotFocus += new RoutedEventHandler(cbox_GotFocus);
                                            cbox.Checked += new RoutedEventHandler(cbox_Checked);
                                            cbox.Unchecked += new RoutedEventHandler(cbox_Unchecked);

                                            _elements.Add(e.templateclauseelementid, cbox);

                                            g2.Children.Add(cbox);


                                        }
                                        else if (er["Element__r_Type__c"].ToString() == "Date")
                                        {
                                            DatePicker dp = new DatePicker();
                                            //cb.Width = 200;
                                            dp.Height = 23;

                                            dp.Name = "cb" + er["Element__r_Name"].ToString().Replace(" ", "");
                                            dp.Margin = new Thickness(120, 4 + (num * 27), 10, 0);
                                            dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;

                                            string dflt = er["Element__r_DefaultValue__c"].ToString();


                                            Element e = new Element();
                                            e.docelementid = "";
                                            e.type = er["Element__r_Type__c"].ToString();
                                            e.controltype = "DatePicker";
                                            e.format = er["Element__r_Format__c"].ToString();
                                            e.templateelementid = er["Element__r_Id"].ToString();
                                            e.conceptid = conceptid; //hold the clause so we know what to open if they click
                                            e.templateclauseelementid = Convert.ToString(er["Id"]); //hold this so we don't update the same field when
                                            e.templateelementname = er["Element__r_Name"].ToString();
                                            e.defaultvalue = dflt;
                                            dp.Tag = e;

                                            dp.GotFocus += new RoutedEventHandler(dp_GotFocus);
                                            dp.SelectedDateChanged += new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);
                                            dp.LostFocus += new RoutedEventHandler(dp_LostFocus);

                                            dp.SelectedDateFormat = DatePickerFormat.Long;

                                            _elements.Add(e.templateclauseelementid, dp);

                                            g2.Children.Add(lbl);
                                            g2.Children.Add(dp);


                                        }
                                        else
                                        {
                                            TextBox tb = new TextBox();
                                            //tb.Width = 200;
                                            tb.Height = 23;
                                            tb.Name = "tb" + er["Element__r_Name"].ToString().Replace(" ", "");
                                            tb.Margin = new Thickness(120, 4 + (num * 27), 10, 0);
                                            tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;

                                            string dflt = er["Element__r_DefaultValue__c"].ToString();

                                            //For now the tag is the ElementId|ClauseElementId|ConceptId|Format - fix this when we have proper objects
                                            //tb.Tag = er["ID"].ToString() + "|" + er["ClauseElementId"].ToString() + "|" + conceptid + "|" + er["Format"].ToString();

                                            if (!_elements.ContainsKey(er["Element__r_Id"].ToString()))
                                            {
                                                Element e = new Element();
                                                e.docelementid = "";
                                                e.type = er["Element__r_Type__c"].ToString();
                                                e.controltype = "TextBox";
                                                e.format = er["Element__r_Format__c"].ToString();
                                                e.templateelementid = er["Element__r_Id"].ToString();
                                                e.conceptid = conceptid; //hold the clause so we know what to open if they click
                                                e.templateclauseelementid = Convert.ToString(er["Id"]); //hold this so we don't update the same field when
                                                e.templateelementname = er["Element__r_Name"].ToString();
                                                e.defaultvalue = dflt;

                                                tb.Tag = e;

                                                if (er["Element__r_Type__c"].ToString() == "Number" || er["Element__r_Type__c"].ToString() == "Currency")
                                                {
                                                    if (er["Element__r_Type__c"].ToString() == "Number") tb.PreviewTextInput += new TextCompositionEventHandler(tb_PreviewTextInput);
                                                    tb.TextAlignment = TextAlignment.Right;

                                                }

                                                // tb.Text = FormatElement(e, dflt);

                                                tb.TextChanged += new TextChangedEventHandler(element_TextChanged);
                                                tb.GotFocus += new RoutedEventHandler(tb_GotFocus);
                                                tb.LostFocus += new RoutedEventHandler(tb_LostFocus);



                                                _elements.Add(e.templateclauseelementid, tb);

                                                g2.Children.Add(lbl);
                                                g2.Children.Add(tb);
                                            }


                                        }

                                        spEl.Children.Add(g2);

                                    }

                                }

                                //Push radio button and its elements
                                spCl.Children.Add(spRb);
                                spCl.Children.Add(spEl);

                                if (Convert.ToString(r["Clause__r_Concept__r_Id"]) != conceptid)
                                {
                                    //Push the last concept
                                    newExp.Content = spCl;
                                    Questions.Children.Add(gExp);

                                    //Create the new expando for this object
                                    conceptid = Convert.ToString(r["Clause__r_Concept__r_Id"]);
                                    conceptname = r["Clause__r_Concept__r_Name"].ToString();

                                    lExp1 = new Button();
                                    lExp1.Style = style;
                                    if (pbInfo == "") lExp1.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                                    lExp1.Margin = new Thickness(0, 8, 10, 0);
                                    lExp1.ToolTip = ConvertHTMLToToolTip(pbInfo);
                                    lExp1.Content = "Info";
                                    lExp1.Height = 28;

                                    pb = new PlaybookTag();
                                    pb.id = conceptid;
                                    pb.html = pbInfo;
                                    pb.type = "Info";

                                    lExp1.Tag = pb;
                                    lExp1.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                                    lExp1.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                    lExp1.Click += new RoutedEventHandler(lExp1_Click);

                                    lExp2 = new Button();
                                    lExp2.Style = style;
                                    if (pbClient == "") lExp2.Foreground = new SolidColorBrush(Color.FromRgb(176, 196, 222));
                                    lExp2.Margin = new Thickness(0, 8, 40, 0);
                                    lExp2.ToolTip = ConvertHTMLToToolTip(pbClient);
                                    lExp2.Content = "Client";

                                    pb = new PlaybookTag();
                                    pb.id = conceptid;
                                    pb.html = pbClient;
                                    pb.type = "Client";
                                    lExp2.Tag = pb;
                                    lExp2.Height = 28;
                                    lExp2.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                                    lExp2.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                    lExp2.Click += new RoutedEventHandler(lExp2_Click);

                                }

                            }

                            // if Allow None is true then add a "None" selection
                            if (Convert.ToBoolean(clauses[0]["Clause__r_Concept__r_AllowNone__c"]))
                            {
                                StackPanel spElNone = new StackPanel();
                                spElNone.Tag = "elementsp";
                                spElNone.Margin = new Thickness(35, 5, 5, 5);
                                StackPanel spRbNone = new StackPanel();
                                spRbNone.Tag = "rbsp";
                                ClauseRadio rb1 = new ClauseRadio("", "None", conceptid, conceptname, num, null, "", "", "", false, "", "");
                                rb1.GroupName = conceptname;
                                rb1.Checked += new RoutedEventHandler(rb1_Checked);
                                rb1.GotFocus += new RoutedEventHandler(rb1_GotFocus);
                                spRbNone.Margin = new Thickness(5, 5, 5, 5);
                                spRbNone.Children.Add(rb1);
                                //Push radio button and its elements
                                spCl.Children.Add(spRbNone);
                                spCl.Children.Add(spElNone);

                                // put in the approval button even though it'll never get triggered for None
                                // Approvals put in the button ---------------------------------------------------
                                Button ApprovalButton = new Button();
                                ApprovalButton.Margin = new Thickness(2, 2, 2, 0);
                                // ApprovalButton.ToolTip = ConvertHTMLToToolTip(pbInfo);
                                ApprovalButton.Content = "Get Approval";
                                ApprovalButton.Tag = conceptid + "|" + conceptname + "|" + "";
                                ApprovalButton.Height = 22;
                                ApprovalButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                                ApprovalButton.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                ApprovalButton.Click += new RoutedEventHandler(ApprovalButton_Click);

                                System.Windows.Controls.Label lbl = new System.Windows.Controls.Label();
                                lbl.Content = "This Clause Requires Approval from: " + "";
                                lbl.Margin = new Thickness(4, 4 + (num * 27), 10, 0);
                                lbl.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                //lbl.Width = 120;

                                Grid gApproval = new Grid();
                                gApproval.Height = 32;
                                gApproval.Children.Add(lbl);
                                gApproval.Children.Add(ApprovalButton);

                                spElNone.Children.Add(gApproval);

                                // just hide
                                gApproval.Visibility = System.Windows.Visibility.Collapsed;

                                //Approvals ---------------------------------------------------


                            }

                            //push the last one
                            newExp.Content = spCl;
                            Questions.Children.Add(gExp);
                        }

                    }
                }

                if (!_attachedmode)
                {
                    // Close the template
                    var docclosetemplate = (Microsoft.Office.Interop.Word._Document)orderdoc;
                    docclosetemplate.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosetemplate);
                }

                // scroll to the top
                this._doc.Characters.First.Select();
            }
            catch (Exception e)
            {
                string message = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) message += " " + e.InnerException.Message;
                MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }
        }




        void unlock_Click(object sender, RoutedEventArgs e)
        {
            // adds a new clause type of clause "free text"
            Button b = (Button)sender;
            string conceptid = Convert.ToString(b.Tag);

            // find the clause
            ClauseRadio rb1 = null;
            foreach (object o in Questions.Children)
            {
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;
                for (int i1 = 0; i1 < spCL.Children.Count; i1++)
                {
                    Object o1 = spCL.Children[i1];
                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {
                        //this is the radiobutton stack panel
                        ClauseRadio rb = (ClauseRadio)sp.Children[0];
                        if (rb.IsChecked == true && rb.conceptid == conceptid)
                        {
                            rb1 = rb;
                            break;
                        }
                    }
                }
            }

            if (rb1 != null)
            {
                if (!rb1.unlock)
                {
                    MessageBoxResult rtn = MessageBox.Show("Are you sure, this will unlock the clause for editing.", "Are you sure?", MessageBoxButton.OKCancel);
                    if (rtn == MessageBoxResult.OK)
                    {
                        //Unlock!

                        //Now unlock the content control in the doc
                        Globals.ThisAddIn.UnlockContractConcept(conceptid, Globals.ThisAddIn.Application.ActiveDocument);

                        //And select
                        Globals.ThisAddIn.SelectConcept(conceptid);

                        //Now mark the button as unlocked
                        Image icon = (Image)b.Content;
                        icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/unlocksmall.png", UriKind.Relative));
                        rb1.unlock = true;
                        b.ToolTip = "Revert Clause back to default and lock";

                        CheckApproval();
                    }
                }
                else
                {
                    MessageBoxResult rtn = MessageBox.Show("Are you sure, this will revert the text to the selected clause and lock for editing", "Are you sure?", MessageBoxButton.OKCancel);
                    if (rtn == MessageBoxResult.OK)
                    {
                        //Lock!

                        SelectClause(rb1);

                        //Update any clauses in the doc with the select clause
                        Globals.ThisAddIn.UpdateContractConcept(rb1.conceptid, rb1.id, rb1.xml, rb1.lastmodified, Globals.ThisAddIn.Application.ActiveDocument, GetElemetValueDict());
                        InitiateElements();

                        //Now mark the button as locked
                        Image icon = (Image)b.Content;
                        icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                        rb1.unlock = true;
                        b.ToolTip = "Unlock Clause";

                        CheckApproval();
                    }
                }
            }

        }

        void ApprovalButton_Click(object sender, RoutedEventArgs e)
        {
            Button b = (Button)sender;
            string[] tag = Convert.ToString(b.Tag).Split('|');
            string conceptid = tag[0];
            string conceptname = tag[1];
            string approver = tag[2];


            SaveAndSendApproval(conceptid, conceptname, approver);
        }


        void dp_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_attachedmode)
            {
                DatePicker dp = (DatePicker)sender;
                Element el = (Element)dp.Tag;

                //Update the doc
                string val = FormatElement(el, dp.Text);
                Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);

                //Upate any other fields
                UpdateElement(el.templateelementid, dp.Text, el.templateclauseelementid);
            }
        }

        void dp_GotFocus(object sender, RoutedEventArgs e)
        {
            if (_attachedmode)
            {
                DatePicker cb = (DatePicker)sender;
                Element el = (Element)cb.Tag;
                Globals.ThisAddIn.SelectConcept(el.conceptid);
            }
        }

        void dp_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_attachedmode)
            {
                //restrict to the values in the list (should make this an option)
                DatePicker dp = (DatePicker)sender;
                Element el = (Element)dp.Tag;
                Globals.ThisAddIn.SelectConcept(el.conceptid);
            }
        }

        void lExp1_Click(object sender, RoutedEventArgs e)
        {
            Button b = (Button)sender;
            PlaybookTag pbt = (PlaybookTag)b.Tag;

            Playbook p = new Playbook();
            p.OpenFromContract(pbt.id, pbt.html, pbt.type);
            p.Show();
        }

        void lExp2_Click(object sender, RoutedEventArgs e)
        {
            Button b = (Button)sender;
            PlaybookTag pbt = (PlaybookTag)b.Tag;

            Playbook p = new Playbook();
            p.OpenFromContract(pbt.id, pbt.html, pbt.type);
            p.Show();
        }



        void element_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_attachedmode)
            {
                TextBox t = (TextBox)sender;
                Element el = (Element)t.Tag;

                //Update the doc
                string val = FormatElement(el, t.Text);
                Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);

                //Upate any other fields
                UpdateElement(el.templateelementid, t.Text, el.templateclauseelementid);
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

        private void InitiateElements()
        {
            //Step through the dictionary of elements and initiate the Content Control in 
            //the doc - also set the value if there is a default
            try
            {
                if (_attachedmode)
                {
                    foreach (string id in _elements.Keys)
                    {
                        FrameworkElement f = _elements[id];
                        Element el = (Element)f.Tag;

                        //Work out the value
                        string val = "";
                        if (el.controltype == "TextBox")
                        {
                            TextBox tb = (TextBox)f;
                            val = tb.Text;

                        }
                        else if (el.controltype == "ComboBox")
                        {
                            ComboBox cb = (ComboBox)f;
                            val = cb.Text;
                        }
                        else if (el.controltype == "CheckBox")
                        {
                            CheckBox cbox = (CheckBox)f;
                            val = cbox.IsChecked.ToString();
                        }
                        else if (el.controltype == "DatePicker")
                        {
                            DatePicker dp = (DatePicker)f;
                            val = dp.Text;
                        }

                        val = FormatElement(el, val);
                        Globals.ThisAddIn.InitiateElement(el.templateelementid, val, el.type, el.format, el.options, el.option1, el.option2);
                    }
                }
            }
            catch (Exception e)
            {
                string message = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) message += " " + e.InnerException.Message;
                MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }

            return;
        }

        private void UpdateElements()
        {
            //Step through the dictionary of elements and update the doc
            try
            {
                foreach (string id in _elements.Keys)
                {
                    FrameworkElement f = _elements[id];
                    Element el = (Element)f.Tag;

                    //Work out the value
                    string val = "";
                    if (el.controltype == "TextBox")
                    {
                        TextBox tb = (TextBox)f;
                        val = tb.Text;
                        val = FormatElement(el, val);
                        tb.Text = val;
                    }
                    else if (el.controltype == "ComboBox")
                    {
                        ComboBox cb = (ComboBox)f;
                        val = cb.Text;
                        val = FormatElement(el, val);
                    }
                    else if (el.controltype == "CheckBox")
                    {
                        CheckBox cbox = (CheckBox)f;
                        val = cbox.IsChecked.ToString();
                        val = FormatElement(el, val);
                    }
                    else if (el.controltype == "DatePicker")
                    {
                        DatePicker dp = (DatePicker)f;
                        val = dp.Text;
                        val = FormatElement(el, val);
                    }

                    if (_attachedmode)
                    {
                        Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);
                    }
                }
            }

            catch (Exception e)
            {
                string message = "Sorry there has been an error - " + e.Message;
                if (e.InnerException != null) message += " " + e.InnerException.Message;
                MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }

            return;
        }


        public void UpdateElement(string elementid, string val, string fromid)
        {
            //Step through the dictionary of elements and update any that match and aren't this one
            //Turn off the events so we don't end up in a big loop!

            foreach (string id in _elements.Keys)
            {
                FrameworkElement f = _elements[id];
                Element el = (Element)f.Tag;

                if (el.templateelementid == elementid && el.templateclauseelementid != fromid)
                {
                    if (el.controltype == "TextBox")
                    {
                        string newval = val;
                        //If its a number then remove any formatting
                        if (el.type.ToLower() == "number")
                        {
                            newval = Regex.Replace(val, "[^0-9.]+", "");
                        }

                        //quick currency handling
                        if (el.type.ToLower() == "currency")
                        {
                            string cur = "";
                            newval = val.Trim();
                            if (newval.ToLower() == "zero")
                            {
                                newval = "0";
                            }
                            else if (newval.Length > 3 && newval.Count(char.IsLetter) >= 3)
                            {
                                cur = newval.Substring(0, 3);
                            }
                            else if (cur == "" && newval.Length > 1)
                            {
                                if (newval.Substring(0, 1) == "$") cur = "USD";
                                if (newval.Substring(0, 1) == "£") cur = "GBP";
                            }

                            newval = Regex.Replace(val, "[^0-9.]+", "");
                            newval = cur + " " + newval;
                        }

                        //Update with Formatting                        
                        val = FormatElement(el, newval);

                        TextBox tb = (TextBox)f;
                        tb.TextChanged -= new TextChangedEventHandler(element_TextChanged);
                        tb.Text = val;
                        tb.TextChanged += new TextChangedEventHandler(element_TextChanged);

                        //Update the doc with the formatting       
                        if (_attachedmode)
                        {
                            Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);
                        }

                    }
                    else if (el.controltype == "ComboBox")
                    {
                        ComboBox cb = (ComboBox)f;
                        cb.SelectionChanged -= new SelectionChangedEventHandler(cb_SelectionChanged);
                        cb.Text = val;
                        cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                    }
                    else if (el.controltype == "CheckBox")
                    {
                        CheckBox cbox = (CheckBox)f;
                        cbox.Checked -= new RoutedEventHandler(cbox_Checked);
                        cbox.Unchecked -= new RoutedEventHandler(cbox_Unchecked);

                        bool bValue;
                        bool isBoolean = bool.TryParse(val, out bValue);
                        if (isBoolean)
                        {
                            cbox.IsChecked = Convert.ToBoolean(bValue);
                        }
                        else
                        {
                            if (val == el.option1)
                            {
                                bValue = true;
                            }
                            else
                            {
                                bValue = false;
                            }
                        }

                        cbox.IsChecked = bValue;

                        cbox.Checked += new RoutedEventHandler(cbox_Checked);
                        cbox.Unchecked += new RoutedEventHandler(cbox_Unchecked);

                    }
                    else if (el.controltype == "DatePicker")
                    {
                        DatePicker dp = (DatePicker)f;
                        dp.SelectedDateChanged -= new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);

                        DateTime dValue;
                        bool isDate = DateTime.TryParse(val, out dValue);
                        if (isDate)
                        {
                            dp.Text = dValue.ToLongDateString();
                        }
                        else
                        {
                            //its not a date! Update back to the current value
                        }

                        //Update with Formatting
                        val = FormatElement(el, dp.Text);
                        if (_attachedmode)
                        {
                            Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);
                        }

                        dp.SelectedDateChanged += new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);

                    }
                }

            }

            return;
        }

        private void LoadElements(string templateelementid, string docelementid, string val)
        {
            //Load the element values and update the element to have the instance id
            string formattedVal = val;
            foreach (string id in _elements.Keys)
            {
                FrameworkElement f = _elements[id];
                Element el = (Element)f.Tag;

                if (el.templateelementid == templateelementid)
                {
                    if (el.controltype == "TextBox")
                    {
                        TextBox tb = (TextBox)f;
                        tb.TextChanged -= new TextChangedEventHandler(element_TextChanged);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        formattedVal = FormatElement(el, val);
                        tb.Text = formattedVal;
                        tb.TextChanged += new TextChangedEventHandler(element_TextChanged);
                    }
                    else if (el.controltype == "ComboBox")
                    {
                        ComboBox cb = (ComboBox)f;
                        cb.SelectionChanged -= new SelectionChangedEventHandler(cb_SelectionChanged);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        formattedVal = FormatElement(el, val);
                        cb.Text = val;
                        cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                    }
                    else if (el.controltype == "CheckBox")
                    {
                        CheckBox cbox = (CheckBox)f;
                        cbox.Checked -= new RoutedEventHandler(cbox_Checked);
                        cbox.Unchecked -= new RoutedEventHandler(cbox_Unchecked);
                        cbox.IsChecked = Convert.ToBoolean(val);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        formattedVal = FormatElement(el, val);
                        cbox.Checked += new RoutedEventHandler(cbox_Checked);
                        cbox.Unchecked += new RoutedEventHandler(cbox_Unchecked);

                    }
                    else if (el.controltype == "DatePicker")
                    {
                        DatePicker dp = (DatePicker)f;
                        dp.SelectedDateChanged -= new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        formattedVal = FormatElement(el, val);
                        dp.Text = val;
                        dp.SelectedDateChanged += new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);
                    }

                    //Update the doc
                    if (_attachedmode)
                    {
                        Globals.ThisAddIn.UpdateElement(el.templateelementid, formattedVal, el.type);
                    }
                }
            }
            return;
        }

        private void LoadElementsFromDoc(string templateelementid, string docelementid, string val)
        {
            bool showrevisions = Globals.ThisAddIn.Application.ActiveDocument.ShowRevisions;
            Globals.ThisAddIn.Application.ActiveDocument.ShowRevisions = false;
            //update the element to have the instance id and get the value from the doc and update the right hand side
            foreach (string id in _elements.Keys)
            {
                FrameworkElement f = _elements[id];
                Element el = (Element)f.Tag;

                //Get the value from the doc
                string formattedVal = Globals.ThisAddIn.GetElementValue(el.templateelementid, el.type);
                string noformattext = RemoveFormatElement(el, formattedVal);

                if (el.templateelementid == templateelementid)
                {
                    if (el.controltype == "TextBox")
                    {
                        TextBox tb = (TextBox)f;
                        tb.TextChanged -= new TextChangedEventHandler(element_TextChanged);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        tb.Text = formattedVal;
                        tb.TextChanged += new TextChangedEventHandler(element_TextChanged);

                        if (val != noformattext)
                        {
                            tb.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFC, 0xDC, 0x3B));
                            tb.ToolTip = "Previous Value:" + FormatElement(el, val);
                        }
                        else
                        {
                            tb.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                            tb.ToolTip = null;
                        }
                    }
                    else if (el.controltype == "ComboBox")
                    {
                        ComboBox cb = (ComboBox)f;
                        cb.SelectionChanged -= new SelectionChangedEventHandler(cb_SelectionChanged);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        cb.Text = formattedVal;
                        cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);

                        if (val != noformattext)
                        {
                            cb.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFC, 0xDC, 0x3B));
                            cb.ToolTip = "Previous Value:" + FormatElement(el, val);
                        }
                        else
                        {
                            cb.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                            cb.ToolTip = null;
                        }
                    }
                    else if (el.controltype == "CheckBox")
                    {
                        CheckBox cbox = (CheckBox)f;
                        cbox.Checked -= new RoutedEventHandler(cbox_Checked);
                        cbox.Unchecked -= new RoutedEventHandler(cbox_Unchecked);
                        cbox.IsChecked = Convert.ToBoolean(noformattext);
                        el.docelementid = docelementid;
                        el.originalvalue = val;
                        cbox.Checked += new RoutedEventHandler(cbox_Checked);
                        cbox.Unchecked += new RoutedEventHandler(cbox_Unchecked);

                    }
                    else if (el.controltype == "DatePicker")
                    {
                        DatePicker dp = (DatePicker)f;
                        dp.SelectedDateChanged -= new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);
                        el.docelementid = docelementid;
                        el.originalvalue = val;

                        string dtval = "";
                        string dtnoformattext = "";
                        try
                        {
                            DateTime dt = Convert.ToDateTime(noformattext);
                            dtnoformattext = dt.ToShortDateString();
                        }
                        catch (Exception)
                        {

                        }

                        dp.Text = noformattext;
                        try
                        {
                            DateTime dt = Convert.ToDateTime(val);
                            dtval = dt.ToShortDateString();
                        }
                        catch (Exception)
                        {

                        }


                        dp.SelectedDateChanged += new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);

                        if (dtval != dtnoformattext)
                        {
                            dp.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFC, 0xDC, 0x3B));
                            dp.ToolTip = "Previous Value:" + FormatElement(el, val);
                        }
                        else
                        {
                            dp.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                            dp.ToolTip = null;
                        }
                    }
                }


            }

            Globals.ThisAddIn.Application.ActiveDocument.ShowRevisions = showrevisions;
            return;
        }


        void rb1_Checked(object sender, RoutedEventArgs e)
        {
            ClauseRadio rb1 = (ClauseRadio)sender;
            //MessageBox.Show("concept>>" + rb1.concept + " > " + rb1.id);

            SelectClause(rb1);

            if (_attachedmode)
            {
                Globals.ThisAddIn.SelectConcept(rb1.conceptid);

                //Update any clauses in the doc with the select clause
                Globals.ThisAddIn.UpdateContractConcept(rb1.conceptid, rb1.id, rb1.xml, rb1.lastmodified, Globals.ThisAddIn.Application.ActiveDocument, GetElemetValueDict());
                InitiateElements();

                //Set the clause to locked
                //find the button
                StackPanel spRb = (StackPanel)rb1.Parent;
                StackPanel spCl = (StackPanel)spRb.Parent;
                Button b = (Button)((Grid)((Expander)spCl.Parent).Parent).Children[3];
                Image icon = (Image)b.Content;
                icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                rb1.unlock = false;
            }

            CheckApproval();
        }


        void rb1_GotFocus(object sender, RoutedEventArgs e)
        {
            if (_attachedmode)
            {
                ClauseRadio rb1 = (ClauseRadio)sender;
                Globals.ThisAddIn.SelectConcept(rb1.conceptid);
            }
        }

        void tb_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox tb = (TextBox)sender;
            Element el = (Element)tb.Tag;

            if (_attachedmode)
            {
                Globals.ThisAddIn.SelectConcept(el.conceptid);
            }

            if (el.type == "Number" || el.type == "Currency")
            {
                //remove formatting
                ((TextBox)sender).Text = RemoveFormatElement(el, ((TextBox)sender).Text);
            }

        }

        void tb_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox tb = (TextBox)sender;
            Element el = (Element)tb.Tag;

            if (_attachedmode)
            {
                Globals.ThisAddIn.SelectConcept(el.conceptid);
            }

            if (el.type == "Number" || el.type == "Currency")
            {
                //add formatting
                ((TextBox)sender).Text = FormatElement(el, ((TextBox)sender).Text);
            }
        }

        //Only allow numbers for Number formated tet boxes
        void tb_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }



        // ok add support for the DefaultSelection__c field - if the clause given 
        // e.g. Matter.Counterparty__c = Test is true THEN select that - select the first
        // one where the default selection is true unless none of them are then pick the 
        // highest order

        void SetDefaultClauses()
        {
            Globals.ThisAddIn.ScreenUpdatingOff();
            foreach (object o in Questions.Children)
            {
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;

                ClauseRadio defaultrb1 = null;
                //int priority = -1;
                int number = -1;

                foreach (object o1 in spCL.Children)
                {

                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {
                        //this is the radiobutton stack panel
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];

                        // get the first one
                        if (number == -1)
                        {
                            defaultrb1 = rb1;
                            number = rb1.number;
                        }

                        // see if we have a default selection formula
                        string dflt = "";
                        if (rb1.dr != null)
                        {
                            dflt = rb1.dr["DefaultSelection__c"].ToString();
                        }

                        if (dflt != "")
                        {
                            // can have multiple selections seperated by | allows basic ORs
                            foreach (string dfltoption in dflt.Split('|'))
                            {
                                // should be in the form Matter.Name = Blah
                                string[] dflta = dfltoption.Split('=');
                                if (dflta.Length == 2)
                                {
                                    string selectfld = dflta[0].Trim();
                                    string selectval = dflta[1].Trim();

                                    string val = this.GetFieldValue(selectfld);

                                    if (selectval == val)
                                    {
                                        defaultrb1 = rb1;
                                    }

                                }
                            }
                        }                        
                    }
                }

                defaultrb1.IsChecked = true;
                SelectClause(defaultrb1);

            }
            CheckApproval();
            Globals.ThisAddIn.ScreenUpdatingOn();

        }


        void SelectClause(ClauseRadio rb1)
        {
            //Find the right panels
            StackPanel spRb = (StackPanel)rb1.Parent;
            StackPanel spCl = (StackPanel)spRb.Parent;
            int els = spCl.Children.IndexOf(spRb);
            StackPanel spEl = (StackPanel)spCl.Children[els + 1];

            //Hide the others
            foreach (Object o in spCl.Children)
            {
                StackPanel sp = (StackPanel)o;
                if ((string)sp.Tag == "elementsp")
                {
                    if (spCl.Children.IndexOf(sp) == els + 1)
                    {
                        sp.Visibility = System.Windows.Visibility.Visible;
                    }
                    else
                    {
                        sp.Visibility = System.Windows.Visibility.Collapsed;
                    }
                }
            }

        }

        private void CheckApproval()
        {
            //Check through all the clauses and see if we need Approval
            bool approval = false;
            foreach (object o in Questions.Children)
            {
                //StackPanel spCL = (StackPanel)((Expander)o).Content;
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;
                approval = false;

                for (int i1 = 0; i1 < spCL.Children.Count; i1++)
                {
                    Object o1 = spCL.Children[i1];
                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {

                        //this is the radiobutton stack panel
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];
                        if (rb1.IsChecked == true)
                        {
                            if (rb1.approver != "")
                            {
                                approval = true;
                            }
                            if (rb1.unlock && rb1.unlockapprover != "")
                            {
                                approval = true;
                            }


                            //Show the approval if we need it
                            StackPanel spRb = (StackPanel)rb1.Parent;
                            StackPanel spCl = (StackPanel)spRb.Parent;
                            int els = spCl.Children.IndexOf(spRb);
                            StackPanel spEl = (StackPanel)spCl.Children[els + 1];
                            Grid appGrid = (Grid)spEl.Children[0];
                            if (approval)
                            {
                                System.Windows.Controls.Label l = (System.Windows.Controls.Label)appGrid.Children[0];
                                l.Content = "This Clause Requires Approval from: " + (rb1.unlock ? rb1.unlockapprover : rb1.approver);
                                appGrid.Visibility = System.Windows.Visibility.Visible;
                            }
                            else
                            {
                                appGrid.Visibility = System.Windows.Visibility.Collapsed;
                            }


                        }
                    }
                }
            }

            if (approval)
            {
                lbApprovals.Visibility = System.Windows.Visibility.Visible;
                btnApprovals.Visibility = System.Windows.Visibility.Visible;
                this.rdTopPanel.Height = new GridLength(110);
                Globals.Ribbons.Ribbon1.Approval(true);
            }
            else
            {
                lbApprovals.Visibility = System.Windows.Visibility.Hidden;
                btnApprovals.Visibility = System.Windows.Visibility.Hidden;
                this.rdTopPanel.Height = new GridLength(85);
                Globals.Ribbons.Ribbon1.Approval(false);
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Word.Selection cs = Globals.ThisAddIn.Application.Selection;

            int start = Globals.ThisAddIn.Application.Selection.Start;
            int end = Globals.ThisAddIn.Application.Selection.End;

            int scroll = Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled;

            //Need to select somewhere editable!
            Globals.ThisAddIn.Application.ActiveDocument.Characters.Last.Select();

            try
            {
                Word.Style s = Globals.ThisAddIn.Application.ActiveDocument.Styles["ContentControl"];
                if (s.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                {
                    s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                }
                else
                {
                    s.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightOrange;
                }
            }
            catch (Exception)
            {
            }

            Globals.ThisAddIn.Application.ActiveDocument.Range(start, end).Select();
            Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = scroll;
        }

        public void Refresh()
        {
            UpdateElements();
        }


        public bool SaveContract(bool ForceSave,bool SaveDoc)
        {
            //Save the Contract    
            Globals.ThisAddIn.RemoveSaveHandler(); // remove the save handler to stop the save calling the save etc.

            //Check we have a name - TODO check name doesn't exist already
            if (tbVersionName.Text == "")
            {
                MessageBox.Show("Please enter a name for the version");
                tbVersionName.Focus();
                return false;
            }

            Globals.ThisAddIn.ProcessingStart("Save Contract");
            DataReturn dr;

            dr = Utility.HandleData(_d.SaveVersion(_versionid, _matterid, _templateid, tbVersionName.Text, tbVersionNumber.Text));
            if (!dr.success) return false;
            _versionid = dr.id;


            if (SaveDoc)
            {
                //Add in the doc id if its not there already
                if (!Globals.ThisAddIn.isUnAttachedContract() && !Globals.ThisAddIn.isContract())
                {
                    Globals.ThisAddIn.AddDocId(_doc, "Contract", _versionid);
                }

                //Save the file as an attachment
                //save this to a scratch file

                Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                string filename = Utility.SaveTempFile(_versionid);
                _doc.SaveAs2(FileName: filename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                //Save a copy!
                Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                string filenamecopy = Utility.SaveTempFile(_versionid + "X");
                Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: false);
                dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                docclose.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                //Now save the file - change this to always save as the version name
                
                Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                string vfilename = this.tbVersionName.Text.Replace(" ", "_") + ".docx";

                if (this._attachmentid==null || this._attachmentid == "")
                {
                    dr = Utility.HandleData(_d.AttachFile(_versionid, vfilename, filenamecopy));
                    _attachmentid = dr.id;
                }
                else
                {
                    dr = Utility.HandleData(_d.UpdateFile(_attachmentid, vfilename, filenamecopy));
                }
            }

            //Go through the Contract Data and Save
            int seq = 1;

            Globals.ThisAddIn.ProcessingUpdate("Get the Clause Values");
            //First the clause selection
            foreach (object o in Questions.Children)
            {
                //StackPanel spCL = (StackPanel)((Expander)o).Content;
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;
                string clauseid = "";
                string conceptid = "";

                //The tag has the docclauseid and the clauseid - get the values out so we know if
                //it is an update and if it is an update if it has changed
                string[] spCLTagA = Convert.ToString(spCL.Tag).Split('|');
                string docclauseid = spCLTagA[0];
                string originalclauseid = spCLTagA.Length == 2 ? spCLTagA[1] : "";

                // if this is a force save then set the docclauseid to blank so new ones get saved
                if (ForceSave)
                {
                    docclauseid = "";
                }


                string docclausename = "";
                string text = "";
                string xml = "";

                for (int i1 = 0; i1 < spCL.Children.Count; i1++)
                {
                    Object o1 = spCL.Children[i1];
                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {

                        //this is the radiobutton stack panel
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];
                        if (rb1.IsChecked == true)
                        {

                            clauseid = rb1.id;
                            conceptid = rb1.conceptid;
                            // name can't be more than 80
                            docclausename = Utility.Truncate(rb1.conceptname, 35) + "-" + Utility.Truncate(rb1.name, 35);

                            // Get the text - can't just pull from the template XML cause it 
                            // will have the element values as well
                            if (_attachedmode)
                            {
                                text = Globals.ThisAddIn.GetContractClauseText(_doc, conceptid);
                                xml = Globals.ThisAddIn.GetContractClauseXML(_doc, conceptid);

                                Globals.ThisAddIn.ProcessingUpdate("Save " + docclausename);

                                //Save to SForce - Clause Selection conceptid has been saved with value clauseid

                                if (docclauseid == "" || clauseid != originalclauseid || rb1.unlock)
                                {
                                    dr = Utility.HandleData(_d.SaveDocumentClause(docclauseid, _versionid, conceptid, clauseid, docclausename, seq++, text, rb1.unlock));
                                    if (!dr.success) return false;

                                    //update ids
                                    docclauseid = dr.id;
                                    spCL.Tag = dr.id + "|" + clauseid;

                                    //if the clause has been unlocked need to save the text to the clause as well
                                    if (rb1.unlock)
                                    {
                                        Globals.ThisAddIn.ProcessingUpdate("Open Scratch");
                                        string clausefilename = Utility.SaveTempFile(docclauseid);

                                        Word.Document scratch = Globals.ThisAddIn.Application.Documents.Add(Visible: false);
                                        scratch.Content.InsertXML(xml);
                                        scratch.SaveAs2(FileName: clausefilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                                        var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                                        docclosescratch.Close(false);
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                                        //Now save the file
                                        Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                                        dr = Utility.HandleData(_d.SaveDocumentFile(docclauseid, clausefilename));
                                    }
                                    else
                                    {
                                        //Remove the attachment if there was one before
                                        //TODO!
                                    }

                                }
                            }
                            else
                            {
                                // Unattached - just save the clause selection

                                Globals.ThisAddIn.ProcessingUpdate("Save " + docclausename);

                                //Save to SForce - Clause Selection conceptid has been saved with value clauseid

                                if (docclauseid == "" || clauseid != originalclauseid)
                                {
                                    dr = Utility.HandleData(_d.SaveDocumentClause(docclauseid, _versionid, conceptid, clauseid, docclausename, seq++, "", false));
                                    if (!dr.success) return false;

                                    //update ids
                                    docclauseid = dr.id;
                                    spCL.Tag = dr.id + "|" + clauseid;

                                }

                            }





                            //Now the elements - just want the ones of the selected clauses
                            StackPanel spEl = (StackPanel)spCL.Children[i1 + 1];
                            if ((string)spEl.Tag == "elementsp")
                            {
                                if (spEl.Children.Count > 0)
                                {
                                    foreach (object o2 in spEl.Children)
                                    {
                                        foreach (object o3 in ((Grid)o2).Children)
                                        {
                                            if (o3.GetType().ToString() == "System.Windows.Controls.TextBox")
                                            {
                                                TextBox tb = (TextBox)o3;
                                                if (tb.Tag.GetType().ToString().EndsWith("Element"))
                                                {
                                                    Element el = (Element)tb.Tag;

                                                    if (ForceSave)
                                                    {
                                                        el.docelementid = "";
                                                        el.originalvalue = "";
                                                    }

                                                    string noformattext = RemoveFormatElement(el, tb.Text);
                                                    string formattedText = FormatElement(el, tb.Text);

                                                    if (el.originalvalue != noformattext)
                                                    {
                                                        dr = Utility.HandleData(_d.SaveDocumentClauseElement(el.docelementid, el.templateelementname, docclauseid, _versionid, el.templateelementid, noformattext, formattedText));
                                                        if (!dr.success) return false;

                                                        //Update the Id and set the orignal value
                                                        el.docelementid = dr.id;
                                                        el.originalvalue = noformattext;
                                                        tb.Tag = el;
                                                        tb.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                                                    }
                                                }
                                            }
                                            else if (o3.GetType().ToString() == "System.Windows.Controls.ComboBox")
                                            {
                                                ComboBox cb = (ComboBox)o3;
                                                if (cb.Tag.GetType().ToString().EndsWith("Element"))
                                                {
                                                    Element el = (Element)cb.Tag;

                                                    if (ForceSave)
                                                    {
                                                        el.docelementid = "";
                                                        el.originalvalue = "";
                                                    }

                                                    string noformattext = cb.Text;
                                                    string formattedText = FormatElement(el, cb.Text);

                                                    if (el.originalvalue != cb.Text)
                                                    {
                                                        dr = Utility.HandleData(_d.SaveDocumentClauseElement(el.docelementid, el.templateelementname, docclauseid, _versionid, el.templateelementid, noformattext, formattedText));
                                                        if (!dr.success) return false;

                                                        //Update the Id
                                                        el.docelementid = dr.id;
                                                        el.originalvalue = noformattext;
                                                        cb.Tag = el;
                                                        cb.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                                                    }
                                                }
                                            }
                                            else if (o3.GetType().ToString() == "System.Windows.Controls.CheckBox")
                                            {
                                                CheckBox cbox = (CheckBox)o3;
                                                if (cbox.Tag.GetType().ToString().EndsWith("Element"))
                                                {
                                                    Element el = (Element)cbox.Tag;

                                                    if (ForceSave)
                                                    {
                                                        el.docelementid = "";
                                                        el.originalvalue = "";
                                                    }

                                                    string noformattext = cbox.IsChecked.ToString();
                                                    string formattedText = FormatElement(el, cbox.IsChecked.ToString());

                                                    if (el.originalvalue != cbox.IsChecked.ToString())
                                                    {
                                                        dr = Utility.HandleData(_d.SaveDocumentClauseElement(el.docelementid, el.templateelementname, docclauseid, _versionid, el.templateelementid, noformattext, formattedText));
                                                        if (!dr.success) return false;

                                                        //Update the Id
                                                        el.docelementid = dr.id;
                                                        el.originalvalue = noformattext;
                                                        cbox.Tag = el;
                                                        cbox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                                                    }
                                                }
                                            }
                                            else if (o3.GetType().ToString() == "System.Windows.Controls.DatePicker")
                                            {
                                                DatePicker dp = (DatePicker)o3;
                                                if (dp.Tag.GetType().ToString().EndsWith("Element"))
                                                {
                                                    Element el = (Element)dp.Tag;

                                                    if (ForceSave)
                                                    {
                                                        el.docelementid = "";
                                                        el.originalvalue = "";
                                                    }

                                                    string noformattext = dp.Text;
                                                    string formattedText = FormatElement(el, dp.Text);
                                                    if (el.originalvalue != dp.Text)
                                                    {
                                                        dr = Utility.HandleData(_d.SaveDocumentClauseElement(el.docelementid, el.templateelementname, docclauseid, _versionid, el.templateelementid, noformattext, formattedText));
                                                        if (!dr.success) return false;

                                                        //Update the Id
                                                        el.docelementid = dr.id;
                                                        el.originalvalue = noformattext;
                                                        dp.Tag = el;
                                                        dp.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFF, 0xFF));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
            }

            Globals.ThisAddIn.AddSaveHandler(); // add it back in
            Globals.ThisAddIn.ProcessingStop("End");
            return true;
        }


        public void LoadContractData(string Id, string DocName)
        {
            //Get all the Clause Data and update the selections
            _versionid = Id;
            this.tbVersionName.Text = DocName;


            //Get all the Clause selections for this Doc
            DataReturn clauseslections;
            clauseslections = Utility.HandleData(_d.GetDocumentClause(_versionid));
            if (clauseslections.success == false) return;

            Globals.ThisAddIn.ProcessingUpdate("Get the Clause Values");
            Globals.ThisAddIn.ScreenUpdatingOff();
            //First the clause selection
            foreach (object o in Questions.Children)
            {

                //Get the conceptId and look up the selection
                string conceptid = Convert.ToString(((Expander)((Grid)o).Children[0]).Tag);
                DataRow[] selection = clauseslections.dt.Select("Concept__c='" + conceptid + "'", "Concept__c");

                //dr = HandleData(_d.GetDocumentClause(_documentid, conceptid));
                //if(!dr.success || dr.dt.Rows.Count==0) return;

                if (selection.Length == 0) return;

                string clauseid = selection[0]["SelectedClause__c"].ToString();
                string docclauseid = selection[0]["Id"].ToString();

                //Now step through the options and pick the right one
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;

                //put the docclauseid on the stack panel tag so it knows to update not create - also put the clauseid so we know if its changed
                spCL.Tag = docclauseid + "|" + clauseid;

                foreach (object o1 in spCL.Children)
                {
                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {
                        //this is the radiobutton stack panel
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];
                        if (clauseid != null && Convert.ToString(clauseid) == rb1.id)
                        {

                            //The template should have the right clause so switch off the hanlder
                            rb1.Checked -= new RoutedEventHandler(rb1_Checked);
                            rb1.IsChecked = true;
                            rb1.Checked += new RoutedEventHandler(rb1_Checked);


                            if (_attachedmode)
                            {
                                //are we unlocked - if so do the right thing
                                if (selection[0]["StandardClause__c"].ToString() == "No")
                                {

                                    //Set the clause to unlocked and update the button
                                    rb1.unlock = true;
                                    StackPanel spRb = (StackPanel)rb1.Parent;
                                    StackPanel spCl = (StackPanel)spRb.Parent;
                                    Button b = (Button)((Grid)((Expander)spCl.Parent).Parent).Children[3];
                                    Image icon = (Image)b.Content;
                                    icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/unlocksmall.png", UriKind.Relative));
                                    b.ToolTip = "Revert Clause back to default and lock";
                                }
                            }

                            //do have to show/hide elements
                            SelectClause(rb1);

                        }

                    }
                }
            }

            //Update Elements - just get them all step through and use update elements
            DataReturn dr = Utility.HandleData(_d.GetDocumentElements(_versionid));
            if (!dr.success) return;
            foreach (DataRow r in dr.dt.Rows)
            {
                //update the value - this will update the doc as well
                LoadElements(r["RibbonElement__c"].ToString(), r["Id"].ToString(), r["Value__c"].ToString());
            }

            Globals.ThisAddIn.ScreenUpdatingOn();
            this._doc.Characters.First.Select();
        }


        public void LoadContractDataFromNegotiatedDoc(string Id, string DocName)
        {
            //Get all the Clause Data and update the selections
            _versionid = Id;
            tbVersionName.Text = DocName;


            //Get all the Clause selections for this Doc
            DataReturn clauseslections;
            clauseslections = Utility.HandleData(_d.GetDocumentClause(_versionid));
            if (!clauseslections.success) return;

            Globals.ThisAddIn.ProcessingUpdate("Get the Clause Values");

            //Create a scratch doc to turn the xml into a range we can work with
            Word.Document scratch = Globals.ThisAddIn.Application.Documents.Add(Visible: false);

            //First the clause selection
            foreach (object o in Questions.Children)
            {

                //Get the conceptId and look up the selection
                string conceptid = Convert.ToString(((Expander)((Grid)o).Children[0]).Tag);
                DataRow[] selection = clauseslections.dt.Select("Concept__c='" + conceptid + "'", "Concept__c");

                if (selection.Length == 0) return;

                string clauseid = selection[0]["SelectedClause__c"].ToString();
                string docclauseid = selection[0]["Id"].ToString();

                //Get the Range from the document and strip out the elements
                Word.Range NegRange = Globals.ThisAddIn.GetContractClauseRange(_doc, conceptid);
                scratch.Range(0, scratch.Content.End).InsertXML(NegRange.WordOpenXML);
                string NegRangeNoElementsText = Utility.RemoveElements(scratch.Content).Text.Trim();

                //Now step through the options and pick the right one
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;

                //put the docclauseid on the stack panel tag so it knows to update not create - also put the clauseid so we know if its changed
                spCL.Tag = docclauseid + "|" + clauseid;

                bool exactmatch = false;

                foreach (object o1 in spCL.Children)
                {
                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {
                        //this is the radiobutton stack panel
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];

                        scratch.Range(0, scratch.Content.End).InsertXML(rb1.xml);
                        string ClauseNoElementsText = Utility.RemoveElements(scratch.Content).Text.Trim();

                        if (NegRangeNoElementsText == ClauseNoElementsText)
                        {
                            exactmatch = true;
                            rb1.Checked -= new RoutedEventHandler(rb1_Checked);
                            rb1.IsChecked = true;
                            rb1.Checked += new RoutedEventHandler(rb1_Checked);

                            //do have to show/hide elements
                            SelectClause(rb1);

                            CheckApproval();
                        }

                    }
                }

                //no exact match then select the last one saved and set the Clause as unlocked
                if (!exactmatch)
                {
                    foreach (object o1 in spCL.Children)
                    {
                        StackPanel sp = (StackPanel)o1;
                        if ((string)sp.Tag == "rbsp")
                        {
                            //this is the radiobutton stack panel
                            ClauseRadio rb1 = (ClauseRadio)sp.Children[0];
                            if (clauseid != null && Convert.ToString(clauseid) == rb1.id)
                            {
                                //The template should have the right clause so switch off the hanlder
                                rb1.Checked -= new RoutedEventHandler(rb1_Checked);
                                rb1.IsChecked = true;
                                rb1.Checked += new RoutedEventHandler(rb1_Checked);

                                //we are unlocked!
                                rb1.unlock = true;

                                StackPanel spRb = (StackPanel)rb1.Parent;
                                StackPanel spCl = (StackPanel)spRb.Parent;
                                Button b = (Button)((Grid)((Expander)spCl.Parent).Parent).Children[3];
                                Image icon = (Image)b.Content;
                                icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/unlocksmall.png", UriKind.Relative));
                                b.ToolTip = "Revert Clause back to default and lock";

                                //do have to show/hide elements
                                SelectClause(rb1);

                                CheckApproval();

                            }

                        }
                    }
                }
            }

            //Update Elements - just get them all step through and use update elements

            DataReturn dr = Utility.HandleData(_d.GetDocumentElements(_versionid));
            if (!dr.success) return;
            foreach (DataRow r in dr.dt.Rows)
            {
                //update the value - this will update the doc as well
                LoadElementsFromDoc(r["RibbonElement__c"].ToString(), r["Id"].ToString(), r["Value__c"].ToString());
            }


            var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
            docclosescratch.Close(false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

        }


        void cb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            ComboBox cb = (ComboBox)sender;
            Element el = (Element)cb.Tag;

            if (cb.SelectedItem != null)
            {
                string val = cb.SelectedItem.ToString();

                //Update the doc
                val = FormatElement(el, val);

                if (_attachedmode)
                {
                    Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);
                }

                //Upate any other fields
                UpdateElement(el.templateelementid, val, el.templateclauseelementid);

                //Update the last selected tag
                el.lastselected = val;
            }
        }




        void cbox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (_attachedmode)
            {

                CheckBox cbox = (CheckBox)sender;
                Element el = (Element)cbox.Tag;
                Globals.ThisAddIn.SelectConcept(el.conceptid);
            }
        }



        void cb_GotFocus(object sender, RoutedEventArgs e)
        {
            if (_attachedmode)
            {
                ComboBox cb = (ComboBox)sender;
                Element el = (Element)cb.Tag;
                Globals.ThisAddIn.SelectConcept(el.conceptid);
            }
        }

        void cb_LostFocus(object sender, RoutedEventArgs e)
        {
            //restrict to the values in the list (should make this an option)
            ComboBox cb = (ComboBox)sender;
            Element el = (Element)cb.Tag;
            if (cb.SelectedItem == null && el.lastselected != "")
            {
                cb.SelectedItem = el.lastselected;
            }

        }

        void cbox_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox cbox = (CheckBox)sender;
            Element el = (Element)cbox.Tag;

            //Update the doc
            string val = FormatElement(el, true.ToString());
            if (_attachedmode)
            {
                Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);
            }

            //Upate any other fields
            UpdateElement(el.templateelementid, true.ToString(), el.templateclauseelementid);

        }

        void cbox_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox cbox = (CheckBox)sender;
            Element el = (Element)cbox.Tag;

            //Update the doc
            string val = FormatElement(el, false.ToString());
            if (_attachedmode)
            {
                Globals.ThisAddIn.UpdateElement(el.templateelementid, val, el.type);
            }

            //Upate any other fields
            UpdateElement(el.templateelementid, false.ToString(), el.templateclauseelementid);
        }


        private string FormatElement(Element e, string value)
        {
            if (e.type.ToLower() == "number")
            {
                if (e.format.ToLower() == "percent")
                {
                    //Work out formatting
                    value = Regex.Replace(value, "[^0-9.]+", "");
                    int num = 0;
                    bool isNumeric = int.TryParse(value, out num);
                    if (isNumeric) value = num.ToString("0") + "%";
                }
                else if (e.format.ToLower() == "percentwords")
                {
                    //Work out formatting
                    value = Regex.Replace(value, "[^0-9.]+", "");
                    int num = 0;
                    bool isNumeric = int.TryParse(value, out num);
                    if (isNumeric) value = Utility.ToText(num).ToLower() + " percent (" + num.ToString("0") + "%)";
                }
                else
                {
                    //Work out formatting
                    value = Regex.Replace(value, "[^0-9.]+", "");
                    int num = 0;
                    bool isNumeric = int.TryParse(value, out num);
                    try
                    {
                        if (isNumeric) value = num.ToString(e.format);
                    }
                    catch (Exception)
                    {
                        //report an error?
                    }
                }
            }
            else if (e.type.ToLower() == "date")
            {

                //Work out formatting
                DateTime dt;
                bool isDate = DateTime.TryParse(value, out dt);
                try
                {
                    if (isDate) value = dt.ToString(e.format);
                }
                catch (Exception)
                {
                    //report an error?
                }


            }
            else if (e.type.ToLower() == "currency")
            {
                //Work out formatting
                string cur = "";
                string newval = value.Trim();
                if (newval.Length > 3 && newval.Count(char.IsLetter) >= 3)
                {
                    cur = newval.Substring(0, 3);
                }
                if (cur == "" && newval.Length > 1)
                {
                    if (newval.Substring(0, 1) == "$") cur = "USD";
                    if (newval.Substring(0, 1) == "£") cur = "GBP";
                }

                newval = Regex.Replace(value, "[^0-9.]+", "");

                int num = 0;
                bool isNumeric = int.TryParse(newval, out num);
                if (num == 0)
                {
                    value = "Zero";
                }
                else
                {
                    value = cur + " " + num.ToString("#,##0") + "";
                }
            }
            else if (e.type.ToLower() == "checkbox")
            {
                //if we have options then use them - if not do True/False with the format
                bool bVal = true;
                bool isBoolean = Boolean.TryParse(value, out bVal);

                if (isBoolean)
                {
                    if (e.option1 != "")
                    {
                        if (bVal)
                        {
                            value = e.option1;
                        }
                        else
                        {
                            value = e.option2;
                            if (value == "") value = " ";
                        }
                    }
                    else
                    {
                        value = bVal.ToString();
                    }

                }

            }
            else
            {
                /* thinking about doing this so I could have the company at the top and the COMPANY in the sign block - 
                 * problem is that need diferent format on the actual instances of the element in the clause - could allow this override the format and store that on the clauseelement
                 * or could just use a formula once we have them =UPPER(PartyAName)
                if (e.format.ToLower() == "lowercase")
                {
                    value = value.ToLower();
                }
                else if (e.format.ToLower() == "uppercase")
                {
                    value = value.ToUpper();
                }
                 * */
            }

            return value;

        }

        private string RemoveFormatElement(Element e, string value)
        {
            if (e.type.ToLower() == "number")
            {
                if (e.format.ToLower() == "percent")
                {
                    value = Regex.Replace(value, "[^0-9.]+", "");
                }
                else if (e.format.ToLower() == "percentwords")
                {
                    //Work out formatting
                    value = Regex.Replace(value, "[^0-9.]+", "");
                }
                else
                {
                    //Work out formatting
                    value = Regex.Replace(value, "[^0-9.]+", "");
                }
            }

            else if (e.type.ToLower() == "currency")
            {
                //Work out formatting
                string cur = "";
                string newval = value.Trim();
                if (newval.ToLower() == "zero")
                {
                    newval = "0";
                }
                else if (newval.Length > 3 && newval.Count(char.IsLetter) >= 3)
                {
                    cur = newval.Substring(0, 3);
                }
                else if (cur == "" && newval.Length > 1)
                {
                    if (newval.Substring(0, 1) == "$") cur = "USD";
                    if (newval.Substring(0, 1) == "£") cur = "GBP";
                }

                newval = Regex.Replace(value, "[^0-9.]+", "");

                if (cur == "") cur = "USD";

                int num = 0;
                bool isNumeric = int.TryParse(newval, out num);
                if (num == 0)
                {
                    value = cur + " 0";
                }
                else
                {
                    value = cur + " " + num.ToString("###0") + "";
                }
            }
            else if (e.type.ToLower() == "checkbox")
            {
                if (e.option1.Trim() == value.Trim())
                {
                    value = "True";
                }
                else
                {
                    value = "False";
                }
            }


            return value;

        }



        public string GetDefaultClauseXML(string conceptid)
        {

            //havne't moved the clauses to the dictionary yet so still stepping though UI
            string xml = "";

            foreach (object o in Questions.Children)
            {
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;

                ClauseRadio defaultrb1 = null;
                //int priority = -1;
                int number = -1;

                foreach (object o1 in spCL.Children)
                {

                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {
                        //this is the radiobutton stack panel - get priority
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];
                        //if (rb1.priority > priority)
                        if (rb1.number > number && rb1.conceptid == conceptid)
                        {
                            defaultrb1 = rb1;
                            //priority = rb1.priority;
                            number = rb1.number;
                        }
                    }
                }
                if (number >= 0)
                {
                    xml = defaultrb1.xml;
                }
            }
            return xml;
        }

        public string ConvertToHTML(string xml)
        {

            //Create a scratch document, save as html, then open the file
            string html = "";

            //OK - there is an oddity in Word - drove me mad! - to Accept changes in the document
            //you have to make the doc ACTIVE - have to! or it gives you a protected error.
            //so get the active document and remember so we can switch back at the end

            Word.Document active = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Document dhtmlcopy2 = null;

            string htmlfilenamecopy = Utility.SaveTempHTMLFile(tbVersionName.Text + "-htmlscratch");

            try
            {

                dhtmlcopy2 = Globals.ThisAddIn.Application.Documents.Add(Visible: false);
                dhtmlcopy2.TrackRevisions = false;
                dhtmlcopy2.TrackFormatting = false;
                dhtmlcopy2.TrackMoves = false;
                dhtmlcopy2.Range().InsertXML(xml);

                dhtmlcopy2.Activate();
                dhtmlcopy2.AcceptAllRevisions();
                Utility.RemoveContentControls(dhtmlcopy2);

            }
            catch (Exception e)
            {

            }


            //TODO
            //switch off the highliting **NOT WORKING!** no time now but save after changing as a word doc and work out why this doesn't work!
            //for now the RemoveContentControl deletes the control and inserts the text away from the control
            //Word.Style s = dhtmlcopy.Styles["ContentControl"];
            //s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;


            dhtmlcopy2.SaveAs2(htmlfilenamecopy, Word.WdSaveFormat.wdFormatFilteredHTML, Encoding: 65001);
            var docclose2 = (Microsoft.Office.Interop.Word._Document)dhtmlcopy2;
            docclose2.Close();

            //Now get the html file
            try
            {
                html = System.IO.File.ReadAllText(htmlfilenamecopy, System.Text.Encoding.UTF8);
                //Clean up the html
                html = html.Replace((char)65533, ' ');
            }
            catch (Exception)
            {
            }

            active.Activate();

            return html;
        }


        public void SaveAndSendApproval(string conceptid, string conceptname, string approver)
        {

            //Check we have a name - TODO check name doesn't exist already
            if (tbVersionName.Text == "")
            {
                MessageBox.Show("Please enter a name for the document");
                tbVersionName.Focus();
                return;
            }

            if (_versionid == "")
            {
                try
                {
                    //Always fails cause the handler returns an error to stop the normal save
                    Globals.ThisAddIn.Application.ActiveDocument.Save();
                }
                catch (Exception)
                {
                }

            }

            Globals.ThisAddIn.ProcessingStart("Save Copy");

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //Have a look for the clause text if we are approving this concept
            string mailsubject = "Approval Required";
            if (approver != "") mailsubject = approver + " " + "Approval Required";
            string mailto = "approver.test@axiomlaw.com";
            string mailbody = "<p style=\"font-family: Calibri; font-size: 14px; color: black\">Please find the attached document for your approval.<br><br><a href='http://www.axiomlaw.com'>Approve Change</a>&nbsp;<a href='http://www.axiomlaw.com'>Deny Change</a><br><br>";

            if (conceptid != "")
            {
                mailbody += "The concept updated is " + conceptname + "<br>";
                string docxml = Globals.ThisAddIn.GetContractClauseXML(_doc, conceptid);
                string html1 = ConvertToHTML(docxml);

                docxml = GetDefaultClauseXML(conceptid);
                string html2 = ConvertToHTML(docxml);



                mailbody += "<br>-------------------------------------------<br>The new text is :<br></p>" + html1;
                mailbody += "<p style=\"font-family: Calibri; font-size: 14px; color: black\"><br>-------------------------------------------<br>The default clause text is :</p><br>" + html2;


            }

            //Save the file as an attachment
            //save this to a scratch file
            //Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
            //string filename = Utility.SaveTempFile(_documentid);               
            //doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatXMLDocument);

            //Save a copy!
            string filenamecopy = Utility.SaveTempFile(tbVersionName.Text);
            Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(doc.FullName, Visible: false);

            //Need to take out the docid!
            Globals.ThisAddIn.DeleteDocId(dcopy);
            Globals.ThisAddIn.AddDocId(dcopy, "ExportContract", _versionid);

            dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

            var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
            docclose.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);



            Outlook.Application o = new Outlook.Application();

            Outlook.MailItem mailItem = o.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = mailsubject;
            mailItem.To = mailto;
            mailItem.HTMLBody = mailbody;
            //mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
            mailItem.Attachments.Add(filenamecopy);
            Globals.ThisAddIn.ProcessingStop("End");


            try
            {
                mailItem.Display(true);
            }
            catch (Exception)
            {
                MessageBox.Show("Please close the active message and try again");
            }


        }


        public void SaveAndSendNeg()
        {

            //Check we have a name - TODO check name doesn't exist already
            if (tbVersionName.Text == "")
            {
                MessageBox.Show("Please enter a name for the document");
                tbVersionName.Focus();
                return;
            }

            // always save
            try
            {
                //Always fails cause the handler returns an error to stop the normal save
                Globals.ThisAddIn.Application.ActiveDocument.Save();
            }
            catch (Exception)
            {
            }



            Globals.ThisAddIn.ProcessingStart("Save Copy");



            //Have a look for the clause text if we are approving this concept
            string mailsubject = tbVersionName.Text + " document to review";
            string mailto = "negotiator@besthedgefunever.com";
            string mailbody = "<Font face='Calibri' size='10px'>Please find the attached document for your approval.<br><br>Please response with any comments.";

            //Save a copy!
            Globals.ThisAddIn.ProcessingUpdate("Save Copy");
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string filenamecopy = Utility.SaveTempFile(tbVersionName.Text);
            Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(doc.FullName, Visible: false);

            //Need to take out the docid!
            Globals.ThisAddIn.DeleteDocId(dcopy);
            Globals.ThisAddIn.AddDocId(dcopy, "ExportContract", _versionid);

            //make the clauses editable
            //Now step through the doc and update the concept if it matches the one we just updated
            object start = dcopy.Content.Start;
            object end = dcopy.Content.End;
            Word.Range r = doc.Range(ref start, ref end);

            // Step through and select the one passed
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                string tag = cc.Tag;
                if (tag != null && tag != "" && cc.Tag.Contains('|'))
                {
                    string[] taga = cc.Tag.Split('|');
                    if (taga[0] == "Concept" || taga[0] == "Element")
                    {
                        cc.LockContents = false;
                    }
                }
            }

            //switch on redlining
            doc.TrackRevisions = true;
            doc.ShowRevisions = true;

            dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);


            var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
            docclose.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

            Outlook.Application o = new Outlook.Application();

            Outlook.MailItem mailItem = o.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = mailsubject;
            mailItem.To = mailto;
            mailItem.HTMLBody = mailbody;
            //mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
            mailItem.Attachments.Add(filenamecopy);
            Globals.ThisAddIn.ProcessingStop("End");


            try
            {
                mailItem.Display(true);
            }
            catch (Exception)
            {
                MessageBox.Show("Please close the active message and try again");
            }


        }

        private void btnApprovals_Click(object sender, RoutedEventArgs e)
        {
            //Dummy Approve!
            Globals.Ribbons.Ribbon1.Approval(false);
            lbApprovals.Visibility = System.Windows.Visibility.Hidden;
            btnApprovals.Visibility = System.Windows.Visibility.Hidden;
            this.rdTopPanel.Height = new GridLength(85);
            Globals.Ribbons.Ribbon1.Approval(false);
        }


        //temp thing to get the element values as a dictionary - should maintain this throughout
        private Dictionary<string, string> GetElemetValueDict()
        {
            Dictionary<string, string> o = new Dictionary<string, string>();
            //Load the element values and update the element to have the instance id
            foreach (string id in _elements.Keys)
            {
                FrameworkElement f = _elements[id];
                Element el = (Element)f.Tag;
                if (!o.ContainsKey(el.templateelementid))
                {
                    if (el.controltype == "TextBox")
                    {
                        TextBox tb = (TextBox)f;
                        o[el.templateelementid] = tb.Text;
                    }
                    else if (el.controltype == "ComboBox")
                    {
                        ComboBox cb = (ComboBox)f;
                        o[el.templateelementid] = cb.Text;
                    }
                    else if (el.controltype == "CheckBox")
                    {
                        CheckBox cbox = (CheckBox)f;
                        o[el.templateelementid] = FormatElement(el, Convert.ToString(cbox.IsChecked));

                    }
                    else if (el.controltype == "DatePicker")
                    {
                        DatePicker dp = (DatePicker)f;
                        o[el.templateelementid] = dp.Text;
                    }
                }


            }
            return o;
        }


        private void btnSavePDF_Click(object sender, RoutedEventArgs e)
        {
            //save this to a scratch file
            Globals.ThisAddIn.ProcessingStart("Save as Pdf");
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //Check we have a name - TODO check name doesn't exist already
            if (tbVersionName.Text == "")
            {
                MessageBox.Show("Please enter a name for the document");
                tbVersionName.Focus();
                return;
            }

            // always save
            try
            {
                //Always fails cause the handler returns an error to stop the normal save
                Globals.ThisAddIn.Application.ActiveDocument.Save();
            }
            catch (Exception)
            {
            }

            //Need to take out the docid!
            Globals.ThisAddIn.DeleteDocId(doc);
            Globals.ThisAddIn.AddDocId(doc, "ExportContract", _versionid);

            // Switch of the element highliting
            // Need to select somewhere editable!
            Globals.ThisAddIn.Application.ActiveDocument.Characters.Last.Select();

            try
            {
                Word.Style s = Globals.ThisAddIn.Application.ActiveDocument.Styles["ContentControl"];
                if (s.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                {
                    s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                }
            }
            catch (Exception)
            {
            }

            // switch on Revisions
            doc.TrackRevisions = true;
            doc.ShowRevisions = true;

            Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
            string filename = Utility.SaveTempFile(_versionid);
            doc.SaveAs2(FileName: filename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

            //Save a copy! - give it the name of the version
            string versionname = tbVersionName.Text;

            Globals.ThisAddIn.ProcessingUpdate("Save PDF");
            string filenamecopy = Utility.SaveTempFile(versionname, "pdf");
            doc.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatPDF);

            var docclose = (Microsoft.Office.Interop.Word._Document)doc;
            docclose.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

            //Now save the file
            Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
            DataReturn dr;
            dr = Utility.HandleData(_d.AttachFile(_versionid, versionname + ".pdf", filenamecopy));
            if (!dr.success) return;

            // open the pdf file in a viewer
            ExternalEditProcess p = new ExternalEditProcess();
            p._id = dr.id;
            p._path = filenamecopy;
            // p._lastwrite = System.IO.File.GetLastWriteTimeUtc(dr.strRtn).ToString();
            // p.EnableRaisingEvents = true;
            // p.Exited += new EventHandler(ExternalEditProcess_HasExited);
            p.StartInfo = new ProcessStartInfo(filenamecopy);
            p.Start();

            Globals.ThisAddIn.ProcessingStop("End");
        }

        private delegate void TwoArgDelegate(String arg1, String arg2);

        private void ExternalEditProcess_HasExited(object sender, System.EventArgs e)
        {
            ExternalEditProcess p = (ExternalEditProcess)sender;

            // just doing pdfs which can't be edited so don't need to worry about this

            //check if the file has been written to 
            /*
            if (System.IO.File.GetLastWriteTimeUtc(p._path).ToString() != p._lastwrite)
            {
                MessageBoxResult rslt = MessageBox.Show("Would you like to update the attachment?", "Update File", MessageBoxButton.OKCancel);
                if (rslt == MessageBoxResult.OK)
                {
                    bsyInd.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new TwoArgDelegate(Update), p._id, p._path);
                }
            }
             * */
        }



        private void btnTemplatePlaybook_Click(object sender, RoutedEventArgs e)
        {
            if (_templateplaybooklink != "")
            {
                // open a browser with the link
                System.Diagnostics.Process.Start(_templateplaybooklink);
            }
        }





        private void cloneBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _cloneBackgroundWorker.DoWork -= (obj, ev) => saveWorkerDoWork(obj, ev);
            _cloneBackgroundWorker.RunWorkerCompleted -= cloneBackgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;

            DataReturn dr = (DataReturn)e.Result;
            AxiomIRISRibbon.Utility.HandleData(dr);

            if (dr.success)
            {
                
                // if this is to be unattached then unattach
                if (this._versioncloneattachedmode)
                {

                    this.tbVersionName.Text = _versionclonename;
                    this.tbVersionNumber.Text = _versionclonenumber;

                    // Save contract will save everything to the new versionid
                    // the true forces the save routine to save the clauses and elements even though they haven't changed
                    this.SaveContract(true,true);
                    // Reload the Data and update the version values
                    this.LoadCompareMenu();
                    this.BuildSidebar();
                    
                }
                else
                {
                    this.tbVersionName.Text = _versionclonename;
                    this.tbVersionNumber.Text = _versionclonenumber;

                    // Save contract will save everything to the new versionid
                    // the true forces the save routine to save the clauses and elements even though they haven't changed
                    this.SaveContract(true, false);
                    // Reload the Data and update the version values
                    this.LoadCompareMenu();
                    this.BuildSidebar();
                    this.UnAttach();
                }


                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
            }


        }


        private void cloneWorkerDoWork(object sender, DoWorkEventArgs e,bool newversionattached)
        {

            // do a save of the current doc
            try
            {
                //Always fails cause the handler returns an error to stop the normal save
                Globals.ThisAddIn.Application.ActiveDocument.Save();
            }
            catch (Exception)
            {
            }

            // now clone!

            DataReturn dr = new DataReturn();
            string newid = "";

            string VersionName = "";
            string VersionNumber = "";
            DataReturn versionmax = _d.GetVersionMax(_matterid);
            string vmax = versionmax.dt.Rows[0][0].ToString();
            double vmaxint = 1;
            if (vmax != null)
            {
                try
                {
                    vmaxint = Convert.ToDouble(vmax) + 1;
                }
                catch (Exception)
                {

                }
            }
            VersionName = "Version " + vmaxint.ToString();
            VersionNumber = vmaxint.ToString();

            _DocumentRow["Name"] = VersionName;
            _DocumentRow["Version_Number__c"] = VersionNumber;

            // if there is an external id field then blank it cause we can't have 2 with the same id
            if (_DocumentRow.Table.Columns.Contains("External_ID__c"))
            {
                _DocumentRow["External_ID__c"] = "";
            }

            DataReturn tempdr = _d.Save(_sDocumentObjectDef, _DocumentRow);
            if (!tempdr.success)
            {
                dr.success = false;
                dr.errormessage += tempdr.errormessage;
            }
            else
            {

                // update the version id to the new one and clear the attachment id                
                newid = tempdr.id;
                _versionid = newid;
                _attachmentid = "";

                _versionclonename = VersionName;
                _versionclonenumber = VersionNumber;
                _versioncloneattachedmode = newversionattached;

            }

            e.Result = dr;

        }


        private void UnAttach()
        {
            // go to Unattach mode! Save the doc as a regular word doc 
            // update the tag to say its unattached and don't do anything
            // if the ribbon buttons or elements are changed


            // if we are importing another doc then do it now
            if (this._versionclonenewdocpath != "")
            {
                // open the new document and get the wordxml and replace it in the current doc

                // NOT SURE ABOUT THIS - think I should just import the doc as is
                // but then I'd have to set it up and open it and then add in the sidebar
                // or something!

                //scratch do to hold the clause 
                Word.Document scratch = Globals.ThisAddIn.Application.Documents.Add(this._versionclonenewdocpath,Visible: false);

                // hide the revisions
                _doc.TrackRevisions = false;
                _doc.ShowRevisions = false;

                // remove content controls
                Globals.ThisAddIn.ProcessingUpdate("Remove Content Controls");
                Globals.ThisAddIn.RemoveContentControls(_doc);

                try
                {
                    _doc.Range().Delete();

                    // delete out the styles! 
                    _doc.Range().set_Style(Word.WdBuiltinStyle.wdStyleNormal);

                    // delete out the pesky tables
                    for (int tablesi = _doc.Range().Tables.Count; tablesi > 0; tablesi--)
                    {
                        _doc.Range().Tables[tablesi].Delete();
                    }

                    _doc.Range().InsertXML(scratch.WordOpenXML);

                } catch(Exception){

                }

                // close the scratch
                var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                docclosescratch.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                this._versionclonenewdocpath = "";
            }


            // take the doc, strip out the content controls and save as a regular word doc
            Globals.ThisAddIn.ProcessingStart("Save as UnAttached Word Doc");

            // remove content controls
            Globals.ThisAddIn.ProcessingUpdate("Remove Content Controls");
            Globals.ThisAddIn.RemoveContentControls(_doc);

            //Need to take out the docid to stop the save handler kicking in
            Globals.ThisAddIn.DeleteDocId(_doc);

            // Switch of the element highliting
            // Need to select somewhere editable!
            _doc.Characters.Last.Select();

            try
            {
                Word.Style s = _doc.Styles["ContentControl"];
                if (s.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                {
                    s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                }
            }
            catch (Exception)
            {
            }

            // switch on Revisions
            _doc.TrackRevisions = true;
            _doc.ShowRevisions = true;

            // Add in the tag
            Globals.ThisAddIn.AddDocId(_doc, "UAContract", _versionid);
            // Set the mode so the save does the right thing
            this.SetAttachedMode("UnAttached");

            // Save - call and the tag will get the handler to save correctly
            try
            {
                //Always fails cause the handler returns an error to stop the normal save
                _doc.Save();
            }
            catch (Exception)
            {
            }

            // Remove the lock buttons now we are unattached
            this.RemoveLockButton();

            Globals.ThisAddIn.ProcessingStop("End");

        }

        private void RemoveLockButton()
        {
            //Check through all the clauses and see if we need Approval            
            foreach (object o in Questions.Children)
            {
                //StackPanel spCL = (StackPanel)((Expander)o).Content;
                StackPanel spCL = (StackPanel)((Expander)((Grid)o).Children[0]).Content;

                for (int i1 = 0; i1 < spCL.Children.Count; i1++)
                {
                    Object o1 = spCL.Children[i1];
                    StackPanel sp = (StackPanel)o1;
                    if ((string)sp.Tag == "rbsp")
                    {
                        // this is the radiobutton stack panel
                        // hide the lock button
                        ClauseRadio rb1 = (ClauseRadio)sp.Children[0];
                        Button b = (Button)((Grid)((Expander)spCL.Parent).Parent).Children[3];
                        b.Visibility = System.Windows.Visibility.Hidden;

                        // move the links to the right
                        Button b1 = (Button)((Grid)((Expander)spCL.Parent).Parent).Children[1];
                        b1.Margin = new Thickness(0, 8, 6, 0);

                        Button b2 = (Button)((Grid)((Expander)spCL.Parent).Parent).Children[2];
                        b2.Margin = new Thickness(0, 8, 36, 0);


                    }
                }
            }

        }

        private void NewVersionContent_ItemClick(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            RadMenuItem mi = e.OriginalSource as RadMenuItem;
     
            string menuheader = mi.Header.ToString();
            string menutag = mi.Tag.ToString();

            // close the menu
            RadContextMenu m = (RadContextMenu)(sender);
            m.IsOpen = false;

            if (menutag == "Template")
            {
                MessageBoxResult rtn = MessageBox.Show("Are you sure, this will create a new version record copying this versions data and attached template contract?", "Are you sure?", MessageBoxButton.OKCancel);
                if (rtn == MessageBoxResult.OK)
                {
                    // New Version
                    if (_versionid != null)
                    {
                        // cloning the version - clear the id and save

                        bsyInd.IsBusy = true;
                        bsyInd.BusyContent = "Saving ...";

                        _DocumentRow["Id"] = "";

                        bool newversionattached = true;

                        _cloneBackgroundWorker = new BackgroundWorker();
                        _cloneBackgroundWorker.DoWork += (obj, ev) => cloneWorkerDoWork(obj, ev, newversionattached);
                        _cloneBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(cloneBackgroundWorker_RunWorkerCompleted);
                        _cloneBackgroundWorker.RunWorkerAsync();
                    }
                }                
            }
            else if (menutag == "UnAttached")
            {
                MessageBoxResult rtn = MessageBox.Show("Are you sure, this will create a new version record copying this versions data but UnAttaching the template contract to a regular word document?", "Are you sure?", MessageBoxButton.OKCancel);
                if (rtn == MessageBoxResult.OK)
                {
                    // clone the version as above - but then unattach the document
                    bsyInd.IsBusy = true;
                    bsyInd.BusyContent = "Saving ...";

                    _DocumentRow["Id"] = "";

                    bool newversionattached = false;

                    _cloneBackgroundWorker = new BackgroundWorker();
                    _cloneBackgroundWorker.DoWork += (obj, ev) => cloneWorkerDoWork(obj, ev, newversionattached);
                    _cloneBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(cloneBackgroundWorker_RunWorkerCompleted);
                    _cloneBackgroundWorker.RunWorkerAsync();

                }                

            }
            else if (menutag == "NewDocument")
            {
                MessageBoxResult rtn = MessageBox.Show("Are you sure, this will create a new version record copying this versions data and will prompt for a new word document to attach as a regular word document?", "Are you sure?", MessageBoxButton.OKCancel);
                if (rtn == MessageBoxResult.OK)
                {
                    // ok ask for the new document ...
                    //Add any file
                    OpenFileDialog dlg = new OpenFileDialog();
                    dlg.Filter = "Word Document (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docx";
                    dlg.Title = "Please select the word doc to attach";

                    Nullable<bool> result = dlg.ShowDialog();

                    // Process open file dialog box results 
                    if (result == true)
                    {
                        // Open document 
                        _versionclonenewdocpath = dlg.FileName;

                        bsyInd.IsBusy = true;
                        bsyInd.BusyContent = "Saving ...";

                        _DocumentRow["Id"] = "";

                        bool newversionattached = false;

                        _cloneBackgroundWorker = new BackgroundWorker();
                        _cloneBackgroundWorker.DoWork += (obj, ev) => cloneWorkerDoWork(obj, ev, newversionattached);
                        _cloneBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(cloneBackgroundWorker_RunWorkerCompleted);
                        _cloneBackgroundWorker.RunWorkerAsync();


                    }



                } 
            }
        }


        private void SetAttachedMode(string AttachedMode)
        {
            if(AttachedMode.ToLower()=="Attached".ToLower() && !_attachedmode){
                _attachedmode = true;
                this.imgAttached.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
                this.imgAttached.ToolTip = "Attached - Document is Locked to Template";

                // hide the attach new version
                this.rmiTemplate.Visibility = System.Windows.Visibility.Visible;


            } else if (AttachedMode.ToLower()=="UnAttached".ToLower() && _attachedmode){
                _attachedmode = false;
                this.imgAttached.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/unlocksmall.png", UriKind.Relative));
                this.imgAttached.ToolTip = "UnAttached - Document is NOT Locked to Template";

                // show the attach new version
                this.rmiTemplate.Visibility = System.Windows.Visibility.Collapsed;

            }
            
        }



        private void ExportContent_ItemClick(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            RadMenuItem mi = e.OriginalSource as RadMenuItem;

            string menuheader = mi.Header.ToString();
            string menutag = mi.Tag.ToString();

            // close the menu
            RadContextMenu m = (RadContextMenu)(sender);
            m.IsOpen = false;

            if (menutag == "Word")
            {
                
                // take the doc, strip out the content controls and save as a regular word doc
                Globals.ThisAddIn.ProcessingStart("Save as Static Word Doc");

                //Check we have a name - TODO check name doesn't exist already
                if (tbVersionName.Text == "")
                {
                    MessageBox.Show("Please enter a name for the document");
                    tbVersionName.Focus();
                    return;
                }

                // always save first
                try
                {
                    //Always fails cause the handler returns an error to stop the normal save
                    Globals.ThisAddIn.Application.ActiveDocument.Save();
                }
                catch (Exception)
                {
                }


                // take out the Id
                Globals.ThisAddIn.DeleteDocId(_doc);

                // create a copy
                string filenamenoext = System.IO.Path.GetFileNameWithoutExtension(_filename);

                // save this to a scratch file
                string filename = AxiomIRISRibbon.Utility.SaveTempFile(filenamenoext);
                _doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatXMLDocument);

                // open it!
                Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: true);

                // if attached then remove content controls
                if (_attachedmode)
                {
                    // remove content controls
                    Globals.ThisAddIn.ProcessingUpdate("Remove Content Controls");

                    object start = newdoc.Content.Start;
                    object end = newdoc.Content.End;
                    Word.Range r = newdoc.Range(ref start, ref end);

                    foreach (Word.ContentControl cc in r.ContentControls)
                    {
                        string tag = cc.Tag;
                        if (tag != null && tag != "" && cc.Tag.Contains('|'))
                        {
                            string[] taga = cc.Tag.Split('|');

                            if (taga.Length > 1 && ((taga[0] == "Concept" && taga[1] != "") || (taga[0] == "Element" && taga[1] != "")))
                            {
                                Word.Range ccr = cc.Range;
                                cc.LockContentControl = false;
                                cc.LockContents = false;
                                cc.Delete(false);
                            }
                        }
                    }

                   
                    // Switch of the element highliting
                    // Need to select somewhere editable!
                    newdoc.Characters.Last.Select();

                    try
                    {
                        Word.Style s = newdoc.Styles["ContentControl"];
                        if (s.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                        {
                            s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                // ** Add back in the doc to the original doc ** so we don't break the original!
                if (_attachedmode)
                {
                    Globals.ThisAddIn.AddDocId(_doc, "Contract", _versionid);
                }
                else
                {
                    Globals.ThisAddIn.AddDocId(_doc, "UAContract", _versionid);
                }

                newdoc.Activate();

                Globals.ThisAddIn.ProcessingStop("End");
            }


            else if (menutag == "PDF")
            {

                // take the doc, strip out the content controls and save as a regular word doc
                Globals.ThisAddIn.ProcessingStart("Save as Static Word Doc");

                //Check we have a name - TODO check name doesn't exist already
                if (tbVersionName.Text == "")
                {
                    MessageBox.Show("Please enter a name for the document");
                    tbVersionName.Focus();
                    return;
                }

                // always save first
                try
                {
                    //Always fails cause the handler returns an error to stop the normal save
                    Globals.ThisAddIn.Application.ActiveDocument.Save();
                }
                catch (Exception)
                {
                }


                // take out the Id
                Globals.ThisAddIn.DeleteDocId(_doc);

                // create a copy
                string filenamenoext = System.IO.Path.GetFileNameWithoutExtension(_filename);

                // save this to a scratch file
                string filename = AxiomIRISRibbon.Utility.SaveTempFile(filenamenoext);
                _doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatXMLDocument);

                // open it!
                Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: true);

                // if attached then remove content controls
                if (_attachedmode)
                {
                    // remove content controls
                    Globals.ThisAddIn.ProcessingUpdate("Remove Content Controls");

                    object start = newdoc.Content.Start;
                    object end = newdoc.Content.End;
                    Word.Range r = newdoc.Range(ref start, ref end);

                    foreach (Word.ContentControl cc in r.ContentControls)
                    {
                        string tag = cc.Tag;
                        if (tag != null && tag != "" && cc.Tag.Contains('|'))
                        {
                            string[] taga = cc.Tag.Split('|');

                            if (taga.Length > 1 && ((taga[0] == "Concept" && taga[1] != "") || (taga[0] == "Element" && taga[1] != "")))
                            {
                                Word.Range ccr = cc.Range;
                                cc.LockContentControl = false;
                                cc.LockContents = false;
                                cc.Delete(false);
                            }
                        }
                    }


                    // Switch of the element highliting
                    // Need to select somewhere editable!
                    newdoc.Characters.Last.Select();

                    try
                    {
                        Word.Style s = newdoc.Styles["ContentControl"];
                        if (s.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                        {
                            s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                newdoc.TrackRevisions = false;
                newdoc.ShowRevisions = false;

                //Save a copy! - give it the name of the version
                string versionname = tbVersionName.Text;

                Globals.ThisAddIn.ProcessingUpdate("Save PDF");
                string filenamecopy = Utility.SaveTempFile(versionname, "pdf");
                newdoc.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatPDF);

                var docclose = (Microsoft.Office.Interop.Word._Document)newdoc;
                docclose.Close(SaveChanges: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                // ** Add back in the doc to the original doc ** so we don't break the original!
                if (_attachedmode)
                {
                    Globals.ThisAddIn.AddDocId(_doc, "Contract", _versionid);
                }
                else
                {
                    Globals.ThisAddIn.AddDocId(_doc, "UAContract", _versionid);
                }

                // open the pdf file in a viewer
                ExternalEditProcess p = new ExternalEditProcess();
                p._id = _versionid;
                p._path = filenamecopy;
                p.StartInfo = new ProcessStartInfo(filenamecopy);
                p.Start();

                Globals.ThisAddIn.ProcessingStop("End");
                
            }
        }



        private void CompareContent_ItemClick(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            RadMenuItem mi = e.OriginalSource as RadMenuItem;

            string menuheader = mi.Header.ToString();
            string menutag = mi.Tag.ToString();

            // close the menu
            RadContextMenu m = (RadContextMenu)(sender);
            m.IsOpen = false;

            // MessageBox.Show("Compare to:" + menuheader);


            // ok the below is the code from the clause comparison
            // I think what we have to do is
            // take the old doc strip out any content controls to get a raw word doc
            // then same for the new one
            // then run the compare and pop that 
            // need to think about the workflow - how do we encorporate changes and stuff
            // but this should show intent!


            // to get the old doc get the version from the menu tag and see if there is one attachment
            string oldversionid = menutag;
            DataReturn dr = Utility.HandleData(this._d.GetVersionAttachments(oldversionid));
            if (!dr.success) return;

            if (dr.dt.Rows.Count == 0)
            {
                MessageBox.Show("Version has no attachment to compare to!");
                return;
            }
            else if (dr.dt.Rows.Count > 1)
            {
                MessageBox.Show("Version has multiple attachments to compare to!");
                return;
            }

            // ok get the old attachment

            string oldattacmentid = dr.dt.Rows[0]["Id"].ToString();
            string oldfilename = dr.dt.Rows[0]["Name"].ToString();
            dr = _d.OpenFile(oldattacmentid, oldfilename);

            Globals.ThisAddIn.RemoveSaveHandler();
            Word.Document olddoc = Globals.ThisAddIn.Application.Documents.Open(dr.strRtn);
            olddoc.Activate();

            this.RemoveControls(olddoc);
            olddoc.TrackRevisions = false;
            olddoc.ShowRevisions = false;
            olddoc.AcceptAllRevisions();

            // save it
            olddoc.Save();

            // now get a copy of the current doc - save it first as the id so we can then open it with the name
            string newscratchfilename = Utility.SaveTempFile(_attachmentid);            
            _doc.SaveAs2(FileName: newscratchfilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

            Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add(newscratchfilename);
            newscratchfilename = Utility.SaveTempFile(_filename);            
            newdoc.SaveAs2(FileName: newscratchfilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

            this.RemoveControls(newdoc);
            newdoc.TrackRevisions = false;
            newdoc.ShowRevisions = false;
            newdoc.AcceptAllRevisions();

            // now compare
            Word.Document compare = Globals.ThisAddIn.Application.CompareDocuments(olddoc, newdoc, Granularity: Word.WdGranularity.wdGranularityCharLevel);
            compare.Activate();

            // close the temp files
            var docclose = (Microsoft.Office.Interop.Word._Document)newdoc;
            docclose.Close(SaveChanges: false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newdoc);

            docclose = (Microsoft.Office.Interop.Word._Document)olddoc;
            docclose.Close(SaveChanges: false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(olddoc);


            Globals.ThisAddIn.AddSaveHandler();
            /* 
            Word.Document oldclause = Globals.ThisAddIn.Application.Documents.Add(Visible: false);
            string oldclausefilename = Utility.SaveTempFile(doc.Name + "-oldclause");
            oldclause.Range().InsertXML(oldxml);

            //get rid of any changes - have to make it the active doc to do this
            oldclause.Activate();   
            oldclause.RejectAllRevisions();

            MakeDropDownElementsText(oldclause);

            //Now update the elements of the scratch
            Utility.UnlockContentControls(scratch);
            UpdateElements(scratch, elementValues);
            //Dropdowns don't diff well (they show as changes, so change the content controls to text - they'll get changed back by initiate)
            MakeDropDownElementsText(scratch);

            //Now run a diff - do it from the old doc rather than a compare so it gives us the redline rather than blue line compare
            string scratchfilename = Utility.SaveTempFile(doc.Name + "-newclause");
            scratch.SaveAs2(FileName: scratchfilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);
            //this is how you do it as a pure compare - Word.Document compare = Application.CompareDocuments(oldclause, scratch,Granularity:Word.WdGranularity.wdGranularityCharLevel);

            oldclause.Compare(scratchfilename, CompareTarget: Word.WdCompareTarget.wdCompareTargetCurrent, AddToRecentFiles: false);
            oldclause.ActiveWindow.Visible = false;

            //Activate the doc - switch of tracking and insert the marked up dif
            doc.Activate();
            doc.TrackRevisions = false;

            // delete out what is there
            cc.Range.Delete();

            // delete out the styles! 
            cc.Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            // delete out the pesky tables
            for (int tablesi = cc.Range.Tables.Count; tablesi > 0; tablesi--)
            {
                cc.Range.Tables[tablesi].Delete();
            }

            cc.Range.FormattedText = oldclause.Content.FormattedText;
            doc.Activate();
            doc.TrackRevisions = true;

            var doccloseoldclause = (Microsoft.Office.Interop.Word._Document)oldclause;
            doccloseoldclause.Close(false);
             
            */


        }


            private void RemoveControls(Word.Document doc){
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');

                        if (taga.Length > 1 && ((taga[0] == "Concept" && taga[1] != "") || (taga[0] == "Element" && taga[1] != "")))
                        {
                            Word.Range ccr = cc.Range;
                            cc.LockContentControl = false;
                            cc.LockContents = false;
                            cc.Delete(false);

                            // *TODO* if none selected we may want to delete the extra return

                        }
                    }
                }

                // Switch of the element highliting
                // Need to select somewhere editable!
                doc.Characters.Last.Select();

                try
                {
                    Word.Style s = doc.Styles["ContentControl"];
                    if (s.Shading.BackgroundPatternColor != Word.WdColor.wdColorAutomatic)
                    {
                        s.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                    }
                }
                catch (Exception)
                {
                }
            }




            private string GetFieldValue(string field)
            {
                string val = "";
                string[] dl = field.Split('.');
                if (dl.Length == 2)
                {
                    if (dl[0] == "Version" || dl[0] == "Version__c")
                    {
                        if (_DocumentRow != null)
                        {
                            if (_DocumentRow.Table.Columns.Contains(dl[1]))
                            {
                                val = _DocumentRow[dl[1]].ToString();
                            }
                        }
                    }
                    else if (dl[0] == "Matter" || dl[0] == "Matter__c")
                    {
                        if (_MatterRow != null)
                        {
                            if (_MatterRow.Table.Columns.Contains(dl[1]))
                            {
                                val = _MatterRow[dl[1]].ToString();
                            }
                        }
                    }
                    else if (dl[0] == "Request" || dl[0] == "Request__c")
                    {
                        if (_RequestRow != null)
                        {
                            if (_RequestRow.Table.Columns.Contains(dl[1]))
                            {
                                val = _RequestRow[dl[1]].ToString();
                            }
                        }

                    }
                 
                }
                return val;
            }



            private void LoadElementsFromDefault()
            {
                try
                {
                    //Load the element values and update the element to have the instance id
                    string formattedVal = "";
                    foreach (string id in _elements.Keys)
                    {
                        FrameworkElement f = _elements[id];
                        Element el = (Element)f.Tag;

                        // get the value
                        string val = "";
                        string dvalue = el.defaultvalue;

                        if (dvalue.StartsWith("="))
                        {
                            string dlookup = dvalue.Substring(1, dvalue.Length - 1);
                            val = GetFieldValue(dlookup);

                            // special formulas
                            if (dvalue.ToLower() == "=Now".ToLower())
                            {
                                val = DateTime.Now.ToLongDateString();
                            }
                        }
                        else
                        {
                            val = dvalue;
                        }


                        if (el.controltype == "TextBox")
                        {

                            TextBox tb = (TextBox)f;
                            tb.TextChanged -= new TextChangedEventHandler(element_TextChanged);
                            el.originalvalue = val;
                            formattedVal = FormatElement(el, val);
                            tb.Text = formattedVal;
                            tb.TextChanged += new TextChangedEventHandler(element_TextChanged);
                        }


                        else if (el.controltype == "ComboBox")
                        {
                            ComboBox cb = (ComboBox)f;
                            cb.SelectionChanged -= new SelectionChangedEventHandler(cb_SelectionChanged);
                            el.originalvalue = val;
                            formattedVal = FormatElement(el, val);
                            cb.Text = val;
                            cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                        }
                        else if (el.controltype == "CheckBox")
                        {
                            CheckBox cbox = (CheckBox)f;
                            cbox.Checked -= new RoutedEventHandler(cbox_Checked);
                            cbox.Unchecked -= new RoutedEventHandler(cbox_Unchecked);
                            cbox.IsChecked = Convert.ToBoolean(val);
                            el.originalvalue = val;
                            formattedVal = FormatElement(el, val);
                            cbox.Checked += new RoutedEventHandler(cbox_Checked);
                            cbox.Unchecked += new RoutedEventHandler(cbox_Unchecked);

                        }
                        else if (el.controltype == "DatePicker")
                        {
                            DatePicker dp = (DatePicker)f;
                            dp.SelectedDateChanged -= new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);
                            el.originalvalue = val;
                            formattedVal = FormatElement(el, dp.Text);
                            dp.Text = val;
                            dp.SelectedDateChanged += new EventHandler<SelectionChangedEventArgs>(dp_SelectedDateChanged);
                        }


                        //Update the doc
                        if (_attachedmode)
                        {
                            Globals.ThisAddIn.UpdateElement(el.templateelementid, formattedVal, el.type);
                        }

                    }
                }
                catch (Exception ex)
                {
                    string message = "Sorry there has been an error - " + ex.Message;
                    if (ex.InnerException != null) message += " " + ex.InnerException.Message;
                    MessageBox.Show(message);
                    // Globals.ThisAddIn.ProcessingStop("Finished");
                }
                return;
            }

            private void btnReset_Click(object sender, RoutedEventArgs e)
            {
                MessageBoxResult rtn = MessageBox.Show("Are you sure, this will reset the Clause selection and Elements to the default values?", "Are you sure?", MessageBoxButton.OKCancel);
                if (rtn == MessageBoxResult.OK)
                {
                    // get the default clauses
                    this.SetDefaultClauses();

                    // populate the elements!
                    this.LoadElementsFromDefault();
                }
            }
    }






    class ExternalEditProcess : Process
    {
        public string _id;
        public string _path;
        public string _lastwrite;
    }

}


