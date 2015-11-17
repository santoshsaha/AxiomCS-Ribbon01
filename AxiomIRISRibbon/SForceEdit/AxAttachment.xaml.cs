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
using System.ComponentModel;
using Telerik.Windows.Controls;
using System.Data;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for AxAttachment.xaml
    /// </summary>
    public partial class AxAttachment : UserControl
    {
        private Data _d;
        private BackgroundWorker _backgroundWorker;
        private bool _gotdata;
        private System.Windows.Media.Color _gbborder;
        public SForceEdit.SObjectDef _sObjectDef;


        public AxAttachment(string ParentType)
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();
            if (StyleManager.ApplicationTheme.ToString() == "Windows8" || StyleManager.ApplicationTheme.ToString() == "Expression_Dark")
            {
                _gbborder = Windows8Palette.Palette.AccentColor;
                //add lines to the grid - windows 8 theme is a bit to white!
                if (StyleManager.ApplicationTheme.ToString() == "Windows8")
                {
                    radGridView1.VerticalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    radGridView1.HorizontalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    radGridView1.GridLinesVisibility = Telerik.Windows.Controls.GridView.GridLinesVisibility.Both;
                }
            }

            SetupAxAttachment(ParentType);
        }

        public void SetupAxAttachment(string ParentType)
        {

            //Attachment has no layout and we arent showing it as a data form anyway
            //so use the Sobject but don;t use AddFields and just build the object here
            //should move to the SobjectDef class as a special case


            _sObjectDef = new SForceEdit.SObjectDef("Attachment");
            _sObjectDef.Parent = "Parent";
            _sObjectDef.ParentType = ParentType;

            //Set up the field in the grid
            List<string> searchCols = new List<string>();
            string settingColumns = Globals.ThisAddIn.GetSettings("Attachment", "Columns");

            if (settingColumns == "")
            {
                settingColumns = "Name|Owner_Name|LastModifiedDate|LastModifiedBy_Name";
            }

            _sObjectDef.GridColumnFields = new List<string>();
            this.radGridView1.Columns.Clear();
            this.radGridView1.AutoGenerateColumns = false;

            sfPartner.DescribeSObjectResult dsr = _d.GetSObject(_sObjectDef.Name);
            _sObjectDef.Label = dsr.label;
            _sObjectDef.PluralLabel = dsr.labelPlural;
            
            // describe object doesn't have the url for attachment, get the url from the binding and add {ID}
            _sObjectDef.Url = _d.GetURL() + "/{ID}";

            _sObjectDef.AddField("Id","","Id","string",true,false,false,null);
            _sObjectDef.AddField("ParentId", "", "ParentId", "string", true, false, false, null);
            _sObjectDef.AddField("ContentType", "", "ContentType", "string", true, false, false, null);

            foreach (string settingCol in settingColumns.Split('|'))
            {
                //Find the Field
                sfPartner.Field f = null;

                string fldname = settingCol;
                if(settingCol.EndsWith("_Name")){
                    fldname = settingCol.Replace("_Name", "") + "Id" ; //dont have to worry about __c to __r can't have custom on attachment
                }

                for (int x = 0; x < dsr.fields.Length; x++)
                {
                    if (dsr.fields[x].name == fldname)
                    {
                        f = dsr.fields[x];
                    }
                }
                
                if (f!=null)
                {

                    //add the field
                    _sObjectDef.AddField(f.name,
                                               f.relationshipName,
                                               f.label,
                                               f.type.ToString(),
                                               true,
                                               f.updateable,
                                               f.createable,
                                               f
                                               );

                    //If this is a Relation add extra field to the field list with the Name
                    if (f.type == sfPartner.fieldType.reference)
                    {
                        _sObjectDef.AddField(f.relationshipName + ".Name",
                            f.relationshipName,
                            f.label, 
                            f.type.ToString(),
                            true,
                            false,
                            false,
                            null);
                    }

                    //keep a list of columns
                    _sObjectDef.GridColumnFields.Add(settingCol);


                    SForceEdit.SObjectDef.FieldGridCol fgc = _sObjectDef.GetField(settingCol);

                    GridViewDataColumn column = new GridViewDataColumn();
                    column.DataMemberBinding = new Binding(fgc.Name);
                    column.Header = fgc.Header;
                    column.UniqueName = fgc.Name;
                    if (fgc.DataType == "date")
                    {
                        column.DataType = typeof(DateTime);
                        column.DataFormatString = "d";
                        column.TextAlignment = TextAlignment.Center;
                    }
                    if (fgc.DataType == "datetime")
                    {
                        column.DataType = typeof(DateTime);
                        column.DataFormatString = "g";
                        column.TextAlignment = TextAlignment.Center;
                    }
                    if (fgc.DataType == "double")
                    {
                        column.DataType = typeof(Double);
                        column.DataFormatString = "0";
                        column.TextAlignment = TextAlignment.Right;
                    }
                    column.MaxWidth = 300;
                    this.radGridView1.Columns.Add(column);
                }
            }
        }


        private void UserControl_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (!_gotdata) LoadData(false);
        }


        private void LoadData(bool justreload)
        {

            //if its just a reload don't scroll back to the top of the grid - keep the current selections
            _sObjectDef.JustReload = justreload;

            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Loading ...";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => WorkerDoWork(obj, ev);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
            _backgroundWorker.RunWorkerAsync();

        }

        public void LoadData(string Id)
        {
            _sObjectDef.ParentId = Id;
            _gotdata = false;
            if (IsVisible) LoadData(false);
        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _backgroundWorker.DoWork -= (obj, ev) => WorkerDoWork(obj, ev);
            _backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;
            DataReturn dr = (DataReturn)e.Result;
            AxiomIRISRibbon.Utility.HandleData(dr);


            // if this is a reload don't pick the first item, the selected item will be selected in a minuite
            // otherwise pick the first one
            if (_sObjectDef.JustReload)
            {
                radGridView1.IsSynchronizedWithCurrentItem = false;
            }
            else
            {
                radGridView1.IsSynchronizedWithCurrentItem = true;
            }
            radGridView1.ItemsSource = dr.dt.DefaultView;

            _gotdata = true;
        }


        void WorkerDoWork(object sender, DoWorkEventArgs e)
        {
            DataReturn dr = _d.GetData(_sObjectDef);
            e.Result = dr;
        }

        private void radGridView1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            //Open The Document
            DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
            EditAttach(r);
        }

        void radGridView1_SelectionChanged(object sender, SelectionChangeEventArgs e)
        {
            e.Handled = true;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            //Add any file
            OpenFileDialog dlg = new OpenFileDialog();

            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;

                //attach
                Attach(filename);                
            }
        }

        private void Attach(string FileName)
        {
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Attaching ...";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => AttachDoWork(obj, ev, FileName);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunAttachCompleted);
            _backgroundWorker.RunWorkerAsync();
        }

        void AttachDoWork(object sender, DoWorkEventArgs e,string FileName)
        {
            DataReturn dr = _d.AttachFile(_sObjectDef.ParentId, System.IO.Path.GetFileName(FileName), FileName);            
            e.Result = dr;
        }


        void backgroundWorker_RunAttachCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _backgroundWorker.DoWork -= (obj, ev) => AttachDoWork(obj, ev,"");
            _backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunAttachCompleted;

            bsyInd.IsBusy = false;
            DataReturn dr = (DataReturn)e.Result;
            if (dr.success)
            {
                //refresh the list
                LoadData(true);
            }
            else
            {
                MessageBox.Show("Sorry there has been a problem:" + dr.errormessage);
            }
        }



        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            //Open The Document
            DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
            EditAttach(r);
        }

        private void EditAttach(DataRow r){
            if(r!=null){

                string filename = r["Name"].ToString();
                if (!filename.Contains("."))
                {
                    //No doc extension - read the content type and add the extension from that
                    if (r["ContentType"].ToString().ToLower().Contains("pdf")) filename += ".pdf";
                    if (r["ContentType"].ToString().ToLower().Contains("msword")) filename += ".doc";
                    if (r["ContentType"].ToString().ToLower().Contains("ms-excel")) filename += ".xls";
                    if (r["ContentType"].ToString().ToLower().Contains("ms-powerpoint")) filename += ".ppt";
                }

                Open(r["Id"].ToString(), filename);
            } 
        }


        private void Open(string Id,string FileName)
        {
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Downloading ...";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => OpenDoWork(obj, ev, Id,FileName);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunOpenCompleted);
            _backgroundWorker.RunWorkerAsync();
        }

        void OpenDoWork(object sender, DoWorkEventArgs e, string Id,string FileName)
        {
            DataReturn dr = _d.OpenFile(Id, FileName);
            e.Result = dr;
        }


        class ExternalEditProcess : Process
        {
            public string _id;
            public string _path;
            public string _lastwrite;
        }

        void backgroundWorker_RunOpenCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _backgroundWorker.DoWork -= (obj, ev) => OpenDoWork(obj, ev, "","");
            _backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunOpenCompleted;

            bsyInd.IsBusy = false;
            DataReturn dr = (DataReturn)e.Result;
            if (dr.success)
            {
                

                if (System.IO.File.Exists(dr.strRtn))
                {
                    
                    //If this is Word then open in a new instance with a sidebar
                    if (dr.strRtn.EndsWith(".docx") || dr.strRtn.EndsWith(".docm") || dr.strRtn.EndsWith(".doc"))
                    {
                        // hide the windows
                        Globals.Ribbons.Ribbon1.CloseWindows();

                        // Word!
                        Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(dr.strRtn);

                        Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                        doc.Activate();                      

                        if (Globals.ThisAddIn.isContract(doc) && (_sObjectDef.ParentType=="Version__c" || _sObjectDef.ParentType=="Document__c"))
                        {
                            Globals.ThisAddIn.AddContractContentControlHandler(doc);
                            //if the action bar isn't open then open it
                            Globals.ThisAddIn.ShowTaskPane(true);

                            ContractEdit.SForceEditSideBar2 u = Globals.ThisAddIn.GetTaskPaneControlContract();
                            if (u != null)
                            {
                                //don't get it to set the defaults if there are going to be values to load
                                u.BuildSideBarFromVersion(_sObjectDef.ParentId, "Attached", dr.id);                                
                            }

                            //Scroll to the top
                            Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = true;
                            Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = 0;
                            
                            // load the data tab
                            u.LoadDataTab(dr.strRtn,_sObjectDef.ParentType, _sObjectDef.ParentId);

                            Globals.ThisAddIn.ProcessingStop("Stop");
                            if (doc != null) doc.Activate();
                            
                        }
                        else if (Globals.ThisAddIn.isUnAttachedContract(doc) && (_sObjectDef.ParentType == "Version__c" || _sObjectDef.ParentType == "Document__c"))
                        {
                            // Unattached - Contract - show the sidebar but just get from the version

                            //if the action bar isn't open then open it
                            Globals.ThisAddIn.ShowTaskPane(true);

                            ContractEdit.SForceEditSideBar2 u = Globals.ThisAddIn.GetTaskPaneControlContract();
                            if (u != null)
                            {
                                //don't get it to set the defaults if there are going to be values to load
                                u.BuildSideBarFromVersion(_sObjectDef.ParentId, "UnAttached", dr.id);
                            }

                            //Scroll to the top
                            Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = true;
                            Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = 0;
                            
                            // load the data tab
                            u.LoadDataTab(dr.strRtn, _sObjectDef.ParentType, _sObjectDef.ParentId);

                            Globals.ThisAddIn.ProcessingStop("Stop");
                            if (doc != null) doc.Activate();
                            
                        }
                        else
                        {
                            Globals.Ribbons.Ribbon1.CloseWindows();

                            // Just a word doc - not attached
                            if (doc != null) doc.Activate();

                            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;                            
                            Globals.ThisAddIn.ShowTaskPaneSFEdit(doc,true, dr.id, dr.strRtn, _sObjectDef.ParentType, _sObjectDef.ParentId);

                            if (doc != null) doc.Activate();

                            // add in a tag so it no where to save
                            Globals.ThisAddIn.AddDocId(doc, "attachmentid", dr.id);

                        }

                    }
                    else
                    {
                        //Open with the system defined editor - catch the process exit and save back to salesforce if it has changed
                        //like outlook does with attachments
                        ExternalEditProcess p = new ExternalEditProcess();
                        p._id = dr.id;
                        p._path = dr.strRtn;
                        p._lastwrite = System.IO.File.GetLastWriteTimeUtc(dr.strRtn).ToString();
                        p.EnableRaisingEvents = true;
                        p.Exited += new EventHandler(ExternalEditProcess_HasExited);
                        p.StartInfo = new ProcessStartInfo(dr.strRtn);
                        p.Start();
                    }
                }
                else
                {
                    MessageBox.Show("Sorry couldn't download the file");
                }
            }
            else
            {
                MessageBox.Show("Sorry there has been a problem:" + dr.errormessage);
            }
        }

        private delegate void TwoArgDelegate(String arg1, String arg2);

        private void ExternalEditProcess_HasExited(object sender, System.EventArgs e)
        {
            ExternalEditProcess p = (ExternalEditProcess)sender;

            //check if the file has been written to 
            if (System.IO.File.GetLastWriteTimeUtc(p._path).ToString() != p._lastwrite)
            {
                MessageBoxResult rslt = MessageBox.Show("Would you like to update the attachment?", "Update File", MessageBoxButton.OKCancel);
                if (rslt == MessageBoxResult.OK)
                {
                    bsyInd.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, new TwoArgDelegate(Update), p._id, p._path);
                }
            }
        }

        private void Update(string Id,string FileName)
        {
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Updating ...";
            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => UpdateDoWork(obj, ev,Id, FileName);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunUpdateCompleted);
            _backgroundWorker.RunWorkerAsync();
        }

        void UpdateDoWork(object sender, DoWorkEventArgs e, string Id,string FileName)
        {
            
            DataReturn dr = _d.UpdateFile(Id,"", FileName);
            e.Result = dr;
        }


        void backgroundWorker_RunUpdateCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _backgroundWorker.DoWork -= (obj, ev) => UpdateDoWork(obj, ev, "","");
            _backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunUpdateCompleted;

            bsyInd.IsBusy = false;
            DataReturn dr = (DataReturn)e.Result;
            if (dr.success)
            {
                //refresh the list
                LoadData(true);
            }
            else
            {
                MessageBox.Show("Sorry there has been a problem:" + dr.errormessage);
            }
        }

        private void Del_Click(object sender, RoutedEventArgs e)
        {
            //Delete the attachment - give them a warning first
            DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
            if (r != null)
            {
                MessageBoxResult rslt = MessageBox.Show("Are you sure you want to delete the attachment:" + r["Name"].ToString() + "?", "Delete File", MessageBoxButton.OKCancel);
                if (rslt == MessageBoxResult.OK)
                {
                    Delete(r["Id"].ToString());
                }
            }     

        }

        private void Delete(string Id)
        {
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Deleting ...";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => DeleteDoWork(obj, ev, Id);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunDeleteCompleted);
            _backgroundWorker.RunWorkerAsync();
        }

        void DeleteDoWork(object sender, DoWorkEventArgs e, string Id)
        {
            DataReturn dr = _d.DeleteFile(Id);
            e.Result = dr;
        }


        void backgroundWorker_RunDeleteCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _backgroundWorker.DoWork -= (obj, ev) => DeleteDoWork(obj, ev, "");
            _backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunDeleteCompleted;

            bsyInd.IsBusy = false;
            DataReturn dr = (DataReturn)e.Result;
            if (dr.success)
            {
                //refresh the list
                LoadData(false);
            }
            else
            {
                MessageBox.Show("Sorry there has been a problem:" + dr.errormessage);
            }
        }

        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
            LoadData(true);
        }

        private void SFButton_Click(object sender, RoutedEventArgs e)
        {
            if (radGridView1.SelectedItem != null)
            {
                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;

                Uri temp = new Uri(_sObjectDef.Url.Replace("{ID}", r["Id"].ToString()));
                string rooturl = temp.Scheme + "://" + temp.Host;

                string frontdoor = rooturl + "/secur/frontdoor.jsp?sid=" + _d.GetSessionId();
                string redirect = frontdoor + "&retURL=" + temp.PathAndQuery;

                System.Diagnostics.Process.Start(redirect);
            }
        }


    }
}
