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

namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for SForceEditSideBar.xaml
    /// </summary>
    public partial class SForceEditSideBar : UserControl
    {

        private Data _d;
        private SForceEdit.SObjectDef _sDocumentObjectDef;
        private SForceEdit.SObjectDef _sMatterObjectDef;
        private SForceEdit.SObjectDef _sRequestObjectDef;
        private SForceEdit.SObjectDef _sActivityObjectDef;
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

        public SForceEditSideBar(string AttachmentId,string FileName,string ParentType,string ParentId)
        {
            _d = Globals.ThisAddIn.getData();

            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _attachmentid = AttachmentId;

            tbDocumentName.Text = System.IO.Path.GetFileName(FileName);
            tbDocumentName.IsReadOnly = true;

            btnSave.IsEnabled = false;

            _filename = System.IO.Path.GetFileName(FileName);
            _parentType = ParentType;
            _parentId = ParentId;

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


            BuildSidebar();

            this.SizeChanged += new SizeChangedEventHandler(Fields_SizeChanged);

            //Wire up the save on the document save - set a setting so we can check this is this doc
            Globals.ThisAddIn.AddDocId(Globals.ThisAddIn.Application.ActiveDocument,"attachmentid",_attachmentid);
            Globals.ThisAddIn.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);

        }

        //Clean up the Save


        ~SForceEditSideBar()
        {
            Globals.ThisAddIn.Application.DocumentBeforeSave -= new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        private void BuildSidebar(){

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

                if (this._d.demoinstance == "general")
                {
                    if (_DocumentRow != null && _DocumentRow.Table.Columns.Contains("Request2__c"))
                    {
                        string id = _DocumentRow["Request2__c"].ToString();
                        _sRequestObjectDef = new SForceEdit.SObjectDef("Request__c");
                        GenerateFields(_sRequestObjectDef);
                        LoadData(id, _sRequestObjectDef);
                    }
                }


            }
            else
            {

                if (_parentType == "Matter__c")
                {
                    _sDocumentObjectDef = new SForceEdit.SObjectDef("Matter__c");
                }
                else
                {
                    //assume ParentType is Version
                    _sDocumentObjectDef = new SForceEdit.SObjectDef("Version__c");
                }

                
                

                // hack for the Document version if the parent is document OR version doesn't exit
                if (_parentType == "Document__c" || _sDocumentObjectDef == null)
                {
                    _sDocumentObjectDef = new SForceEdit.SObjectDef("Document__c");
                    GenerateFields(_sDocumentObjectDef);
                    LoadData(_parentId, _sDocumentObjectDef);
                }
                else
                {

                    //Console.WriteLine(_parentType);
                    //_sDocumentObjectDef = new SForceEdit.SObjectDef("Version__c");
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
                        //AddGrid(_sActivityObjectDef);
                        // GenerateFields(_sActivityObjectDef);
                        //  LoadData(id, _sActivityObjectDef);

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
            }
        }

        private void GenerateFields(SForceEdit.SObjectDef sObj)
        {
            

            sfPartner.DescribeSObjectResult dsr = _d.GetSObject(sObj.Name);
            sObj.Label = dsr.label;
            sObj.PluralLabel = dsr.labelPlural;
            if (_setgbborder) sObj.SetGBBorder(_gbborder);

            sObj.BuildCompactLayouts(_d, FieldChanged,SalesforcePressed,OpenPressed);

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
            if (r!=null && r.Table.Columns.IndexOf("RecordTypeId") >= 0)
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
                if(!sObj.RecordTypeMapping.ContainsKey(sObj.DefaultRecordType)) rid = sObj.RecordTypeMapping.ElementAt(0).Key;
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
                _DocumentRow = r;
                SForceEdit.Utility.UpdateForm(FindStackPanel("Document__c"), _DocumentRow);
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
                        if (g.Children[j].GetType() == typeof(StackPanel)) { ((StackPanel)g.Children[j]).Width = width; }

                        /* everything is wrapped in a stackpanel now so just need to resize it
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
                         * */
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
            Word.Document  doc = Globals.ThisAddIn.Application.ActiveDocument;

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
                if (s1.Tag!=null && s1.Tag.ToString() == sObjName)
                {
                    return s1;
                }
            }
            return ret;
        }



        void FieldChanged()
        {
            StackPanel flds;
            bool changes=false;
            //Document
            if (_DocumentRow != null)
            {
                if (this._d.demoinstance == "general" || this._d.demoinstance == "isda")
                {
                    flds = FindStackPanel("Document__c");
                    _DocumentRow.BeginEdit();
                    if (SForceEdit.Utility.UpdateRow(flds, _DocumentRow)) changes = true;
                    _DocumentRow.CancelEdit();
                }
                else
                {
                    flds = FindStackPanel("Version__c");
                    _DocumentRow.BeginEdit();
                    if (SForceEdit.Utility.UpdateRow(flds, _DocumentRow)) changes = true;
                    _DocumentRow.CancelEdit();
                }
            }

            //Matter
            if (_MatterRow != null)
            {
                flds = FindStackPanel("Matter__c");
                _MatterRow.BeginEdit();
                if (SForceEdit.Utility.UpdateRow(flds, _MatterRow)) changes = true ;
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
            if (sObjectType == "Document__c" && _RequestRow != null)
            {
                temp = new Uri(_sRequestObjectDef.Url.Replace("{ID}", _RequestRow["Id"].ToString()));
            }

            if (temp!=null)
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
                Edit ed = new Edit("Version__c", _DocumentRow["Id"].ToString());
                ed.Show();                
            }
            if (sObjectType == "Matter__c" && _MatterRow != null)
            {
                Edit ed = new Edit("Matter__c", _MatterRow["Id"].ToString());
                ed.Show(); 
            }
            if (sObjectType == "Request__c" && _RequestRow != null)
            {
                Edit ed = new Edit("Request__c", _RequestRow["Id"].ToString());                
                ed.Show();
            }

        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Saving ...";

            if (this._d.demoinstance == "general" || this._d.demoinstance == "isda") {
                _DocumentChanges = SForceEdit.Utility.UpdateRow(FindStackPanel("Document__c"), _DocumentRow);
            }
            else
            {
                _DocumentChanges = SForceEdit.Utility.UpdateRow(FindStackPanel("Version__c"), _DocumentRow);
            }
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
            if (_DocumentRow != null)
            {
                if (this._d.demoinstance == "general" || this._d.demoinstance == "isda")
                {
                    SForceEdit.Utility.UpdateForm(FindStackPanel("Document__c"), _DocumentRow);
                }
                else
                {
                    SForceEdit.Utility.UpdateForm(FindStackPanel("Version__c"), _DocumentRow);
                }
            }
            if(_MatterRow!=null) SForceEdit.Utility.UpdateForm(FindStackPanel("Matter__c"), _MatterRow);
            if (_RequestRow != null) SForceEdit.Utility.UpdateForm(FindStackPanel("Request__c"), _RequestRow);


            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }

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

        

    }
}

