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

namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for AxObject.xaml
    /// </summary>
    public partial class AxObject : UserControl
    {
        private Data _d;

        private BackgroundWorker _backgroundWorker;
        private BackgroundWorker _saveBackgroundWorker;
        private BackgroundWorker _buttonBackgroundWorker;

        public SForceEdit.SObjectDef _sObjectDef;

        private System.Windows.Media.Color _gbborder;
        private bool _setgbborder = false;

        private string _CurrentRecordTypeId;
        private string _CurrentRecordTypeName;
        
        private List<SForceEdit.AxObject> _childobj;
        private SForceEdit.AxAttachment _childattachments;

        private bool _gotdata;
        private string _selectidonceloaded;

        private bool _rootobject;
        private bool _paged;
        private bool _filter;
        private bool _search;
        private bool _zoom;

        public RadWindow _rootWindow;        
        public AxObject _parentObject;

        struct Breadcrumb {
            public DataRow _r;
            public string _objectname;
        }

        public AxObject(string sObject,RadWindow r){

            InitAxObject();
            _rootWindow = r;
            _zoom = false;
            SetupAxObject(sObject, "",null,"","");

            //bread crumbs starts collapsed
            bread1.Visibility = System.Windows.Visibility.Collapsed;
            BreadGridRow.Height = new GridLength(0);
        }

        //Open with an specific object opened
        public AxObject(string sObject, RadWindow r,String Id)
        {

            InitAxObject();
            _rootWindow = r;
            _zoom = false;
            SetupAxObject(sObject, "", null,Id,"");

            //bread crumbs starts collapsed
            bread1.Visibility = System.Windows.Visibility.Collapsed;
            BreadGridRow.Height = new GridLength(0);

        }

        // Zooom!
        public AxObject(string Mode,string sObject, RadWindow r,String Id)
        {

            InitAxObject();
            _rootWindow = r;
            _zoom = true;
           
            SetupAxObject(sObject, "", null,Id,"");

            //remove the grid!
            this.split1.Visibility = System.Windows.Visibility.Hidden;
            this.split1.Width = 0;
            this.split1.Margin = new Thickness(-10);
            r.Width = 600;

            //bread crumbs starts collapsed
            bread1.Visibility = System.Windows.Visibility.Collapsed;
            BreadGridRow.Height = new GridLength(0);

        }

        public AxObject(string sObject, string Relation,AxObject parent,string tabfilters)
        {
            _zoom = false;
            InitAxObject();
            _parentObject = parent;
            SetupAxObject(sObject, Relation, null,"",tabfilters);
        }

        private void InitAxObject()
        {           
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);
            //ScrollViewer not getting set by set theme just set manually for now
            StyleManager.SetTheme(FieldContent, StyleManager.ApplicationTheme);
            radGridView1.SelectionChanged += new EventHandler<SelectionChangeEventArgs>(radGridView1_SelectionChanged);

            _d = Globals.ThisAddIn.getData();
            if (StyleManager.ApplicationTheme.ToString() == "Windows8" || StyleManager.ApplicationTheme.ToString() == "Expression_Dark")
            {
                _setgbborder = true;
                _gbborder = Windows8Palette.Palette.AccentColor;
                //add lines to the grid - windows 8 theme is a bit to white!
                if (StyleManager.ApplicationTheme.ToString() == "Windows8")
                {
                    radGridView1.VerticalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    radGridView1.HorizontalGridLinesBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1"));
                    radGridView1.GridLinesVisibility = Telerik.Windows.Controls.GridView.GridLinesVisibility.Both;

                    this.tbSearchButton.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D6D4D4"));
                }
            }

           
        }

        public void SetupAxObject(string sObject,string Relation,DataRow r,string id,string tabfilters)
        {
            _gotdata = false;

            if (Relation == "")
            {
                _rootobject = true;
                _paged = true; _filter = true; _search = true;
            }
            else
            {
                _rootobject = false;
                _paged = false; _filter = false; _search = false;

                if (tabfilters != "") _filter = true;
            }

            _CurrentRecordTypeId = "";
            // cbFilter.SelectedIndex = 0;
            OpenSObject(sObject,Relation,tabfilters);

            //if this is the root then switch on the paging/filtering/searching - this may be set more fine-grained in settings
            if (_rootobject)
            {                
                split1.InitialPosition = Telerik.Windows.Controls.Docking.DockState.DockedLeft; //root has list to the left 
                split1.Width = 350;
                rp2.PaneHeaderVisibility = System.Windows.Visibility.Visible;
            }
            else
            {               
                split1.InitialPosition = Telerik.Windows.Controls.Docking.DockState.DockedTop; //sub tabs have list to the top                
                split1.Height = 200;
                rp2.PaneHeaderVisibility = System.Windows.Visibility.Collapsed;
                bread1.Visibility = System.Windows.Visibility.Collapsed;
                BreadGridRow.Height = new GridLength(0);
            }
            

            if (_search)
            {
                //wire up the Search Clear Command
                tbSearch.ClearCommand = new DelegateCommand(x =>
                {
                    if (_sObjectDef.Search != "")
                    {
                        _sObjectDef.Search = "";
                        LoadData(false);
                    }
                    tbSearch.Value = "";
                });
            }

            if (r == null)
            {
                //set the id if this is being loaded with an id
                _sObjectDef.Id = id;
                //problem with the thread - probably should be doing all of the above a background thread
                if (_rootobject || _zoom) LoadDataDirect();
            }
            else
            {
                // list view - no filter and no search
                _filter = false;
                if (tabfilters != "") _filter = true;
                _search = false;
                //load from the calling double click - update the grid and the panel
                radGridView1.ItemsSource = r.Table.DefaultView;
                foreach (DataRowView rw in (DataView)radGridView1.ItemsSource)
                {
                    if (rw["Id"].ToString() == r["Id"].ToString()) radGridView1.SelectedItem = rw;
                }
                
                LoadRow(r);
            }



        }

        public void OpenSObject(string sObject,string Relation,string TabFilters)
        {

            _sObjectDef = new SForceEdit.SObjectDef(sObject);

            // set the tab filters to overide the filters if they are set
            if (TabFilters != "") _sObjectDef.TabFilters = TabFilters;

            if(_paged) _sObjectDef.AddPaging();            
            if(_filter) _sObjectDef.Filter = cbFilter.Text;

            AddFields();

            _sObjectDef.Parent = Relation;
        }

        void radGridView1_SelectionChanged(object sender, SelectionChangeEventArgs e)
        {
            if (radGridView1.SelectedItem != null)
            {
                
                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
                LoadRow(r);
                this._sObjectDef.JustReload = false;
                e.Handled = true;
            }
        }

        public string GetLabelPlural()
        {
            return _sObjectDef.PluralLabel;
        }

        private void LoadRow(DataRow r){

            // Update Header
            if (r.Table.Columns.Contains(_sObjectDef.NameField)) rpDetail1.Header = _sObjectDef.Label + " > " + r[_sObjectDef.NameField].ToString();

            //If we have record types then update the form layout if it has changed
            if (r.Table.Columns.IndexOf("RecordTypeId") >= 0)
            {
                string rid = r["RecordTypeId"].ToString();
                if (rid != "" && rid != _CurrentRecordTypeId)
                {

                    sfPartner.RecordTypeMapping m = _sObjectDef.RecordTypeMapping[rid];

                    StackPanel sp = _sObjectDef.Layouts[m.layoutId];
                    FieldContent.Content = sp;
                    _CurrentRecordTypeId = rid; //m.layoutId;
                    _CurrentRecordTypeName = m.name;

                    //UpdatePickLists(m);

                    DisplayButtons();

                    UpdateTextWidth();
                }
            }
           
            //Update subpanels
            foreach (SForceEdit.AxObject axObj in _childobj)
            {
                axObj.LoadData(r["Id"].ToString(), r[_sObjectDef.NameField].ToString(), this._sObjectDef.JustReload);
            }
            if (_childattachments!=null)
            {
                _childattachments.LoadData(r["Id"].ToString());
            }

            StackPanel flds = (StackPanel)FieldContent.Content;
            if (flds != null) SForceEdit.Utility.UpdateForm(flds, r);
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }

        private void ClearRow()
        {
            //Clear the Form and the subtabs
            rpDetail1.Header = _sObjectDef.Label;
            foreach (SForceEdit.AxObject axObj in _childobj)
            {
                axObj.LoadData("","",false);
            }
            if (_childattachments != null)
            {
                _childattachments.LoadData("");
            }

            StackPanel flds = (StackPanel)FieldContent.Content;
            if (flds != null) SForceEdit.Utility.UpdateForm(flds, null);
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;

        }


        private void AddFields()
        {
            //Add Fields! - get the Layout definition
            Fields.Children.Clear();

            if (_setgbborder) _sObjectDef.SetGBBorder(_gbborder);

            //This is where all the work is done - reads the layouts and builds everything
            _sObjectDef.BuildLayouts(_d,FieldChanged);
            
            if (_rootobject && _rootWindow!=null)
            {
                if (!_zoom)
                {
                    _rootWindow.Header = "Axiom IRIS - " + _sObjectDef.PluralLabel;
                }
                else
                {
                    _rootWindow.Header = "" + _sObjectDef.PluralLabel;                                        
                }
            }

            //Remove any sub tabs
            for (int x = tab1.Items.Count-1; x >0; x--)
            {
                tab1.Items.RemoveAt(x);
            }

            //Sort out the Grid - add all the columns
            this.radGridView1.Columns.Clear();
            this.radGridView1.AutoGenerateColumns = false;
            _sObjectDef.AddColumns(radGridView1);


            //Switch on and off features
            if (_paged && rp2.Header.ToString()=="Search")
            {
                radDataPager1.Visibility = System.Windows.Visibility.Visible;                
            }
            else
            {
                radDataPager1.Visibility = System.Windows.Visibility.Collapsed;
            }

            if (rp2.Header.ToString() == "List")
            {
                _filter = false;
            }


            if (!_filter)
            {
                cbFilter.Visibility = System.Windows.Visibility.Collapsed;
            }
            else
            {
                cbFilter.Visibility = System.Windows.Visibility.Visible;
            }


            // sort out the filter entries
            if (_filter)
            {                
                cbFilter.Items.Clear();
                if (_sObjectDef.GridFilters.Count == 0)
                {
                    // none defined just add in My Records and All Records
                    RadComboBoxItem rci = new RadComboBoxItem();
                    rci.Content = "My Records";
                    cbFilter.Items.Add(rci);
                    rci = new RadComboBoxItem();
                    rci.Content = "All Records";
                    cbFilter.Items.Add(rci);
                    cbFilter.SelectedIndex = 0;
                    _sObjectDef.Filter = "My Records";
                }
                else
                {
                    // add the filters
                    string dflt = "";
                    string first = "";
                    foreach (string key in _sObjectDef.GridFilters.Keys)
                    {
                        SObjectDef.FilterEntry f = _sObjectDef.GridFilters[key];

                        RadComboBoxItem rci = new RadComboBoxItem();
                        rci.Content = f.Name;
                        cbFilter.Items.Add(rci);

                        if (f.Default)
                        {
                            dflt = f.Name;
                        }

                        if (first == "")
                        {
                            first = f.Name;
                        }

                    }

                    if (dflt != "")
                    {
                        cbFilter.Text = dflt;
                        _sObjectDef.Filter = dflt;
                    }
                    else
                    {
                        cbFilter.Text = first;
                        _sObjectDef.Filter = first;
                    }
                }

                // Add in the handler
                cbFilter.SelectionChanged += cbFilter_SelectionChanged;
            }



            // Set the form to have the default Record Type
            if (_sObjectDef.DefaultRecordType == null)
            {
                string firstkey = _sObjectDef.Layouts.ElementAt(0).Key;
                StackPanel sp = _sObjectDef.Layouts[firstkey];
                FieldContent.Content = sp;
                _CurrentRecordTypeId = firstkey;
                _CurrentRecordTypeName = firstkey;
            }
            else
            {
                StackPanel sp = _sObjectDef.Layouts[_sObjectDef.DefaultRecordType];
                FieldContent.Content = sp;
                _CurrentRecordTypeId = _sObjectDef.DefaultRecordType;
                _CurrentRecordTypeName = _sObjectDef.DefaultRecordType;
            }

            if (_sObjectDef.RecordTypes.Keys.Count > 0)
            {
                foreach (string rt in _sObjectDef.RecordTypes.Keys)
                {
                    RadMenuItem mi = new RadMenuItem();
                    mi.Header = _sObjectDef.RecordTypes[rt];
                    mi.Tag = rt;
                    this.NewButtonContent.Items.Add(mi);
                }
            }
            else
            {
                RadMenuItem mi = new RadMenuItem();
                mi.Header = _sObjectDef.Label;
                mi.Tag = "";
                this.NewButtonContent.Items.Add(mi);
            }

            // Show the buttons
            DisplayButtons();

            //Wire Up the ReSize
            rpg1.SizeChanged += new SizeChangedEventHandler(Fields_SizeChanged);


            //Subtabs
            _childobj = new List<SForceEdit.AxObject>();
            _childattachments = null;
            if (_rootobject)
            {

                string tabdefinition = Globals.ThisAddIn.GetSettings(_sObjectDef.Name, "Tabs");
                if (tabdefinition != "")
                {
                    foreach (string tdef in tabdefinition.Split('|'))
                    {

                        if (tdef == "Attachment" || tdef == "Attachment:")
                        {
                            //Add an attacment object - just a grid with buttons - give it the parent its attached to so it knows
                            //what to do if the attachment is opened in a sidebar
                            _childattachments = new SForceEdit.AxAttachment(_sObjectDef.Name);
                            RadTabItem tabItem = new RadTabItem()
                            {
                                Header = "Attachments",
                                Content = _childattachments
                            };
                            tab1.Items.Add(tabItem);
                        }

                        else if (tdef.Contains(':'))
                        {
                            string subobject = tdef.Split(':')[0];
                            string rel = tdef.Split(':')[1];
                            //Add the subpanel if we have the definition

                            if (_d.GetSObject(subobject) != null)
                            {
                                // get the tab filters if they exist
                                string tabfilters = Globals.ThisAddIn.GetSettings(_sObjectDef.Name, "TabFilters", subobject);

                                SForceEdit.AxObject axObj = new SForceEdit.AxObject(subobject, rel, this, tabfilters);
                                _childobj.Add(axObj);
                                RadTabItem tabItem = new RadTabItem()
                                {
                                    Header = axObj.GetLabelPlural(),
                                    Content = axObj
                                };
                                tab1.Items.Add(tabItem);
                            }
                        }



                    }

                }
            }
            else
            {
                //If not root add the attachments if it is defined
                string tabdefinition = Globals.ThisAddIn.GetSettings(_sObjectDef.Name, "Tabs");
                if (tabdefinition != "")
                {
                    foreach (string tdef in tabdefinition.Split('|'))
                    {
                        if (tdef == "Attachment" || tdef == "Attachment:")
                        {
                            //Add an attacment object - just a grid with buttons - give it the parent its attached to so it knows
                            //what to do if the attachment is opened in a sidebar
                            _childattachments = new SForceEdit.AxAttachment(_sObjectDef.Name);
                            RadTabItem tabItem = new RadTabItem()
                            {
                                Header = "Attachments",
                                Content = _childattachments
                            };
                            tab1.Items.Add(tabItem);
                        }
                    }
                }

            }

            this.UpdateTextWidth();
        }


        // Display the buttons depending on the Record Type - this should be done in
        // the Sobject the same way as the fields (I think) but just doing it here for now
        void DisplayButtons()
        {
            // Remove any buttons
            for (int x = this.tbObjectButtons.Items.Count - 1; x >= 0; x--)
            {
                if (this.tbObjectButtons.Items[x].GetType() == typeof(RadButton))
                {
                    RadButton b = (RadButton)this.tbObjectButtons.Items[x];
                    if (b.Tag != null && b.Tag.ToString().StartsWith("Custom*"))
                    {
                        this.tbObjectButtons.Items.RemoveAt(x);
                    }
                }
            }
            for (int x = this.tbDataObjectButtons.Items.Count - 1; x >= 0; x--)
            {
                if (this.tbDataObjectButtons.Items[x].GetType() == typeof(RadButton))
                {
                    RadButton b = (RadButton)this.tbDataObjectButtons.Items[x];
                    if (b.Tag != null && b.Tag.ToString().StartsWith("Custom*"))
                    {
                        this.tbDataObjectButtons.Items.RemoveAt(x);
                    }
                }
            }

            // Get the button definition
            string buttondefinition = Globals.ThisAddIn.GetSettings(_sObjectDef.Name, "Buttons");
            bool hasDataButtons = false;

            if (buttondefinition != "")
            {

                foreach (string bdef in buttondefinition.Split('|'))
                {

                    string[] b = bdef.Split(':');
                    string name = b[0];
                    string type = (b.Length > 1) ? b[1] : "";
                    string action = (b.Length > 2) ? b[2] : "";
                    string recordtypes = (b.Length > 3) ? b[3] : "";
                    string confirm = (b.Length > 4) ? b[4] : "";


                    if (type == "Add" || type == "Data")
                    {

                        RadButton rB = new RadButton()
                        {
                            Content = name,
                            BorderThickness = new Thickness(0),
                            Height = 22,
                            Margin = new Thickness(3, 0, 3, 0),
                            Tag = "Custom*" + bdef
                        };
                        rB.Click += rB_Click;

                        if (type == "Add")
                        {
                            this.tbObjectButtons.Items.Add(rB);
                        }
                        else
                        {
                            bool showbutton = true;
                            if (recordtypes.Trim() != "")
                            {
                                showbutton = false;
                                string[] rt = recordtypes.Split(',');
                                foreach (string r in rt)
                                {
                                    if (r == _CurrentRecordTypeName)
                                    {
                                        showbutton = true;
                                    }
                                }
                            }
                            if (showbutton)
                            {
                                hasDataButtons = true;
                                this.tbDataObjectButtons.Items.Add(rB);
                            }
                        }
                    }
                    else if (type == "AddSeperator" || type == "DataSeperator")
                    {
                        RadToolBarSeparator tbS = new RadToolBarSeparator();
                        if (type == "AddSeperator")
                        {
                            this.tbObjectButtons.Items.Add(tbS);
                        }
                        else
                        {
                            bool showbutton = true;
                            if (recordtypes.Trim() != "")
                            {
                                showbutton = false;
                                string[] rt = recordtypes.Split(',');
                                foreach (string r in rt)
                                {
                                    if (r == _CurrentRecordTypeName)
                                    {
                                        showbutton = true;
                                    }
                                }
                            }
                            if (showbutton)
                            {
                                hasDataButtons = true;
                                this.tbDataObjectButtons.Items.Add(tbS);
                            }
                        }
                    }
                }
            }
            if (hasDataButtons)
            {
                if (this.tbDataObjectButtons.Visibility == System.Windows.Visibility.Collapsed)
                {
                    this.tbDataObjectButtons.Visibility = System.Windows.Visibility.Visible;
                    rowDataButtons.Height = new GridLength(28);
                    this.tbDataObjectButtons.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, EmptyDelegate);
                }

            }
            else
            {
                if (this.tbDataObjectButtons.Visibility == System.Windows.Visibility.Visible)
                {
                    this.tbDataObjectButtons.Visibility = System.Windows.Visibility.Collapsed;
                    rowDataButtons.Height = new GridLength(0);
                    this.tbDataObjectButtons.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, EmptyDelegate);
                }

            }
        }


        void Fields_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateTextWidth();
        }

        void UpdateTextWidth()
        {


            //Resize all the children
            
            StackPanel CurrentSPFields = (StackPanel)FieldContent.Content;
            for (int i = 0; i < CurrentSPFields.Children.Count; i++)
            {

                if (CurrentSPFields.Children[i].GetType() == typeof(Telerik.Windows.Controls.GroupBox))
                {
                    Telerik.Windows.Controls.GroupBox gb = (Telerik.Windows.Controls.GroupBox)CurrentSPFields.Children[i];
                    Grid g = (Grid)gb.Content;

                    int cols = g.ColumnDefinitions.Count;
                    double width = (rpg1.ActualWidth / (cols / 2)) - 150;
                    if (width < 100) width = 100;

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


        // go though each of the PickLists and remove any of the entries that aren't available
        // for this record type mapping
        void UpdatePickLists(sfPartner.RecordTypeMapping rtm)
        {
            StackPanel CurrentSPFields = (StackPanel)FieldContent.Content;
            for (int i = 0; i < CurrentSPFields.Children.Count; i++)
            {

                if (CurrentSPFields.Children[i].GetType() == typeof(Telerik.Windows.Controls.GroupBox))
                {
                    Telerik.Windows.Controls.GroupBox gb = (Telerik.Windows.Controls.GroupBox)CurrentSPFields.Children[i];
                    Grid g = (Grid)gb.Content;

                    for (int j = 0; j < g.Children.Count; j++)
                    {
                        if (g.Children[j].GetType() == typeof(StackPanel)) { 
                        StackPanel spcontrol = (StackPanel)g.Children[j];
                        foreach (Control spchildControl in spcontrol.Children)
                        {
                            Control childControl = spchildControl;

                            if (childControl.GetType() == typeof(RadComboBox))
                            {

                                Telerik.Windows.Controls.RadComboBox cb = (Telerik.Windows.Controls.RadComboBox)childControl;
                                if (cb.Tag != null)
                                {
                                    string[] n = cb.Tag.ToString().Split('|');
                                    string name = n[0];
                                    if (name != "RecordTypeId")
                                    {
                                        foreach (sfPartner.PicklistForRecordType prt in rtm.picklistsForRecordType)
                                        {
                                            if (prt.picklistName == name)
                                            {
                                                // ok - step through the items in the picklist and see if they are here
                                                for (int z = 0; z < cb.Items.Count; z++)
                                                {
                                                    if (cb.Items[z].GetType() == typeof(string))
                                                    {

                                                        bool found = false;
                                                        foreach (sfPartner.PicklistEntry entry in prt.picklistValues)
                                                        {
                                                            if (entry.value == cb.Items[z].ToString())
                                                            {
                                                                found = true;
                                                            }
                                                        }
                                                        if (!found)
                                                        {
                                                            cb.Items.RemoveAt(z);
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
                }
            }            
        }


        void FieldChanged()
        {
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _backgroundWorker.DoWork -= (obj, ev) => WorkerDoWork(obj, ev);
            _backgroundWorker.RunWorkerCompleted -= backgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;
            DataReturn dr = (DataReturn)e.Result;
            AxiomIRISRibbon.Utility.HandleData(dr);

            // if this is a reload don't pick the first item, the selected item will be selected in a minuite
            radGridView1.IsSynchronizedWithCurrentItem = false;
            radGridView1.ItemsSource = dr.dt.DefaultView;

            if (_sObjectDef.Paging)
            {
                //if this is not jsut a reload then scroll to the top of the grid
                if (!_sObjectDef.JustReload) radGridView1.ScrollIndexIntoView(0);

                radDataPager1.PageIndex = _sObjectDef.CurrnetPage;
                radDataPager1.PageSize = _sObjectDef.RecordsPerPage;
                radDataPager1.ItemCount = _sObjectDef.TotalRecords;

                if (_sObjectDef.SortColumn != "")
                {
                    radGridView1.Columns[_sObjectDef.SortColumn].SortingState = _sObjectDef.SortDir == "ASC" ? SortingState.Ascending : SortingState.Descending;
                }
            }

            // Select the Row - if no data then clear the form
            if (dr.dt.Rows.Count == 0)
            {
                ClearRow();
            }
            else
            {

                // if a button has crated a new item then select that 
                if (_selectidonceloaded != null && _selectidonceloaded != "")
                {

                    // select the row that has been returned
                    bool found = false;
                    foreach (DataRowView rw in (DataView)this.radGridView1.ItemsSource)
                    {
                        if (rw["Id"].ToString() == _selectidonceloaded)
                        {
                            radGridView1.SelectedItem = rw;
                            radGridView1.ScrollIntoViewAsync(rw,null);
                            found = true;
                        }
                    }
                    if (!found) ClearRow();

                    _selectidonceloaded = "";

                }
                else if (_sObjectDef.JustReload && _sObjectDef.SelectedId!="")
                {
                    bool found = false;
                    foreach (DataRowView rw in (DataView)radGridView1.ItemsSource)
                    {                        
                        if (rw["Id"].ToString() == _sObjectDef.SelectedId)
                        {
                            radGridView1.SelectedItem = rw;
                            found = true;
                        }                        
                    }
                    if (!found) ClearRow();
                }
                else
                {
                    radGridView1.SelectedItem = radGridView1.Items[0];
                  // radGridView1.SelectedItem = ((DataView)(radGridView1.ItemsSource)).Table.Rows[0];
                }



            }
            

            _sObjectDef.SelectedId = "";
            _gotdata = true;
        }


        void WorkerDoWork(object sender, DoWorkEventArgs e)
        {
            DataReturn dr = _d.GetData(_sObjectDef);
            e.Result = dr;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {

            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Saving ...";

            DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
            StackPanel flds = (StackPanel)FieldContent.Content;

            // check all required fields are filled in
            string message = SForceEdit.Utility.CheckRequireFieldsHaveValues(flds, this._sObjectDef);
            if (message != "")
            {
                bsyInd.IsBusy = false;
                MessageBox.Show("Please fill in values for the required fields:" + message);
                return;
            }

            //Update the row
            r.BeginEdit();
            SForceEdit.Utility.UpdateRow(flds, r);

            //save
            _saveBackgroundWorker = new BackgroundWorker();
            _saveBackgroundWorker.DoWork += (obj, ev) => saveWorkerDoWork(obj, ev, r);
            _saveBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(saveBackgroundWorker_RunWorkerCompleted);
            _saveBackgroundWorker.RunWorkerAsync();

        }

        void saveBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _saveBackgroundWorker.DoWork -= (obj, ev) => saveWorkerDoWork(obj, ev, null);
            _saveBackgroundWorker.RunWorkerCompleted -= saveBackgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;

            DataReturn dr = (DataReturn)e.Result;
            AxiomIRISRibbon.Utility.HandleData(dr);

            DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
            if (dr.success)
            {                
                if (r["Id"].ToString() == "")
                {
                    // it is an add - set to the returned id
                    r["Id"] = dr.id;
                }

                r.EndEdit();
                r.AcceptChanges();

                // reload the saved line to update for triggers etc.
                // update the parent if there is one, incase the parent has
                // been updated
                _selectidonceloaded = r["id"].ToString();
                if (_parentObject != null)
                {
                    _parentObject.LoadData(true);
                    this.LoadData(true);
                }
                else
                {
                    this.LoadData(true);
                }
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;

            }
            else
            {
                r.CancelEdit();
            }

        }


        void saveWorkerDoWork(object sender, DoWorkEventArgs e, DataRow r)
        {
            DataReturn dr = _d.Save(_sObjectDef, r);
            e.Result = dr;
        }


        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //revert
            StackPanel flds = (StackPanel)FieldContent.Content;
            if (radGridView1.SelectedItem != null)
            {
                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
                if (r["Id"].ToString() != "")
                {
                    SForceEdit.Utility.UpdateForm(flds,r);
                } else {
                    r.Delete();
                    SForceEdit.Utility.UpdateForm(flds, null);
                }
            }

            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }



        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
            LoadData(true);
        }

        private void radDataPager1_PageIndexChanged(object sender, PageIndexChangedEventArgs e)
        {
            if (e.OldPageIndex >= 0)
            {
                _sObjectDef.CurrnetPage = e.NewPageIndex;
                LoadData(false);
            }
        }

        private void cbFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_sObjectDef != null)
            {
                _sObjectDef.Filter = cbFilter.Text;
                LoadData(false);
            }
        }


        //Load data direct - load the data on the same thread, used by root node when the window is initialised
        public void LoadDataDirect(){

            // if this is a zoom then switch off the Filter
            // Id will be passed
            if (_zoom)
            {
                _sObjectDef.Filter = null;
            }

            DataReturn dr = _d.GetData(_sObjectDef);
            AxiomIRISRibbon.Utility.HandleData(dr);

            radGridView1.IsSynchronizedWithCurrentItem = true;
            radGridView1.ItemsSource = dr.dt.DefaultView;

            if (_sObjectDef.Paging)
            {
                radDataPager1.PageIndex = _sObjectDef.CurrnetPage;
                radDataPager1.PageSize = _sObjectDef.RecordsPerPage;
                radDataPager1.ItemCount = _sObjectDef.TotalRecords;

                if (_sObjectDef.SortColumn != "")
                {
                    radGridView1.Columns[_sObjectDef.SortColumn].SortingState = _sObjectDef.SortDir == "ASC" ? SortingState.Ascending : SortingState.Descending;
                }
            }

            _gotdata = true;
            
        }





        private void LoadData(bool justreload)
        {

            // remember the selected id
            _sObjectDef.SelectedId = "";
            if (radGridView1.SelectedItem != null)
            {
                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
                _sObjectDef.SelectedId = r["Id"].ToString();
            }

            if (!justreload)
            {
                foreach (SForceEdit.AxObject axObj in _childobj)
                {
                    // clear any selection on the children
                    axObj.radGridView1.SelectedItem = null;
                }
            }


            //if its just a reload don't scroll back to the top of the grid - keep the current selections
            _sObjectDef.JustReload = justreload;

            //clear the id if that has been set
            if (!justreload) _sObjectDef.Id = "";

            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Loading ...";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => WorkerDoWork(obj, ev);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
            _backgroundWorker.RunWorkerAsync();

        }

        public void LoadDataZoom(string Id){
            _sObjectDef.Id = Id;
            bsyInd.IsBusy = true;
            bsyInd.BusyContent = "Loading ...";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += (obj, ev) => WorkerDoWork(obj, ev);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
            _backgroundWorker.RunWorkerAsync();
        }

        public void LoadData(string Id,string Name,bool justreload)
        {
            _sObjectDef.ParentId = Id;
            _sObjectDef.ParentName = Name;
            _gotdata = false;
            if (IsVisible) LoadData(justreload);
        }


        private void radGridView1_Sorting(object sender, GridViewSortingEventArgs e)
        {
            if (_sObjectDef.Paging && _sObjectDef.TotalRecords > _sObjectDef.RecordsPerPage)
            {
                e.Cancel = true;

                //Have to get the Query name of the field - this is what actually goes in the OrderBy (this will be diferent if its a relation eg OWNER_NAME and OWNER.NAME)
                //need the name as well so we can show the sort indicator on the column once we have the databack                
                string qField = _sObjectDef.GetField(e.Column.UniqueName).QueryField;

                if (e.OldSortingState == SortingState.None)
                {
                    _sObjectDef.SortColumn = e.Column.UniqueName;
                    _sObjectDef.SortQueryField = qField;
                    _sObjectDef.SortDir = "ASC";
                }
                else if (e.OldSortingState == SortingState.Ascending)
                {
                    _sObjectDef.SortColumn = e.Column.UniqueName;
                    _sObjectDef.SortQueryField = qField;
                    _sObjectDef.SortDir = "DESC";
                }
                else
                {

                    _sObjectDef.SortColumn = "";
                    _sObjectDef.SortQueryField = "";
                    _sObjectDef.SortDir = "";
                }

                _sObjectDef.CurrnetPage = 0;
                LoadData(false);
            }
        }


        private void tbSearchButton_Click(object sender, RoutedEventArgs e)
        {
            if (tbSearch.Text != _sObjectDef.Search)
            {
                _sObjectDef.Search = tbSearch.Text;
                LoadData(false);
            }
        }


        private void tbSearch_ValueChanged(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            // when the search is cleared then reload
            if (tbSearch.Text == "")
            {
                _sObjectDef.Search = tbSearch.Text;
                LoadData(false);
            }
        }

        private void tbSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                _sObjectDef.Search = tbSearch.Text;
                LoadData(false);
            }
        }




        private void SFButton_Click(object sender, RoutedEventArgs e)
        {
            if(radGridView1.SelectedItem!=null){
                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;

                Uri temp = new Uri(_sObjectDef.Url.Replace("{ID}", r["Id"].ToString()));
                string rooturl = temp.Scheme + "://" + temp.Host;
                
                string frontdoor = rooturl + "/secur/frontdoor.jsp?sid=" + _d.GetSessionId();
                string redirect = frontdoor + "&retURL=" + temp.PathAndQuery;

                System.Diagnostics.Process.Start(redirect);
            }
        }

        private void UserControl_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (!_rootobject && !_gotdata) LoadData(false);
        }


        private void radDocking_PaneStateChange(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            RadPane pane = (RadPane)e.OriginalSource;
            if (pane != null && !pane.IsPinned)
            {

                Telerik.Windows.Controls.Docking.AutoHideArea area = pane.Parent as Telerik.Windows.Controls.Docking.AutoHideArea;
                if (area != null && area.HasItems)
                {
                    RadPane selectedPane = area.SelectedPane;
                    if (selectedPane != null)
                    {
                        selectedPane.MouseUp += new MouseButtonEventHandler(selectedPane_MouseUp);
                    }
                }
            }
        }

        void selectedPane_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var pane = sender as RadPane;
            if (pane != null)
            {
                pane.IsPinned = true;
                pane.MouseUp -= selectedPane_MouseUp;
            }
        }

        private void radGridView1_MouseDoubleClick(object sender, MouseButtonEventArgs e)        
        {

            // switch this off for now - too hard to explain - might have to come up with a 
            //   diferent way to do this
            /*
            FrameworkElement originalSender = e.OriginalSource as FrameworkElement;
            if (originalSender != null)
            {               
                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
                if (r != null)
                {
                    //MessageBox.Show("The double-clicked row is object:" +_sObjectDef.Name + " id:" +  r["Id"] );
                    if (!_rootobject)
                    {
                        //Open this as the root object and add the breadcrumb                        
                        _parentObject.AddBreadcrumb();
                        _parentObject.SetupAxObject(_sObjectDef.Name, "", r,"","");                        
                    }

                }
            }
             * */
        }

        private static System.Action EmptyDelegate = delegate() { };

        public void AddBreadcrumb(){

            if (radGridView1.SelectedItem != null)
            {
                //if its not shown show it
                if (bread1.Visibility == System.Windows.Visibility.Collapsed)
                {
                    bread1.Visibility = System.Windows.Visibility.Visible;
                    BreadGridRow.Height = new GridLength(28);
                    bread1.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, EmptyDelegate);
                    ShowListView();
                } 

                DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
                if (r.Table.Columns.Contains(_sObjectDef.NameField))
                {

                    Breadcrumb b = new Breadcrumb();
                    b._objectname = _sObjectDef.Name;
                    b._r = r;

                    Telerik.Windows.Controls.Label l;
                    if (bread1.Children.Count > 0)
                    {
                        l = new Telerik.Windows.Controls.Label();
                        l.Content = ">";
                        bread1.Children.Add(l);
                    }

                    l = new Telerik.Windows.Controls.Label();
                    l.Content = _sObjectDef.Label + ":" + r[_sObjectDef.NameField];
                    l.MouseEnter += new MouseEventHandler(breadcrumb_l_MouseEnter);
                    l.MouseLeave += new MouseEventHandler(breadcrumb_l_MouseLeave);
                    l.MouseDown += new MouseButtonEventHandler(breadcrumb_l_MouseDown);
                    l.Tag = b;
                    bread1.Children.Add(l);
                }
            }
        }


        public void ShowSearchView()
        {
            rp2.Header = "Search";
            searchbarrow.Height = new GridLength(28);
            //searchbarrow.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, EmptyDelegate);
          //  searchbar1.Visibility = System.Windows.Visibility.Visible;
          //  searchbar2.Visibility = System.Windows.Visibility.Visible;
        }

        public void ShowListView()
        {
            //Change the header of the search to list and hide 
            rp2.Header = "List";
            searchbarrow.Height = new GridLength(0);
            searchbarrow.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, EmptyDelegate);
         //   searchbar1.Visibility = System.Windows.Visibility.Hidden;
          //  searchbar2.Visibility = System.Windows.Visibility.Hidden;
        }


        public void ClearBreadcrumb()
        {
            //remove the breadcrumbs
            bread1.Children.Clear();

            //collapse the bar
            bread1.Visibility = System.Windows.Visibility.Collapsed;
            BreadGridRow.Height = new GridLength(0);

            //Show the search view
            ShowSearchView();
        }


        void breadcrumb_l_MouseDown(object sender, MouseButtonEventArgs e)
        {
            int x = -1;
            Telerik.Windows.Controls.Label thisone = (Telerik.Windows.Controls.Label)sender;
            Breadcrumb b = (Breadcrumb)thisone.Tag;

            for (int i = bread1.Children.Count - 1; i >= 0; i--)
            {
                Telerik.Windows.Controls.Label thatone = (Telerik.Windows.Controls.Label)bread1.Children[i];
                if (thatone == thisone)
                {
                    x = i;
                }
            }

            //remove the rest to the right and the spacer to the left
            if (x > -1)
            {
                for (int i = bread1.Children.Count - 1; i >= (x==0?0:x-1); i--)
                {
                    bread1.Children.RemoveAt(i);
                }
            }

            //if no breadcrumbs hide
            if (x == 0)
            {
                ClearBreadcrumb();
            }

            //now open the breadcrumb
            SetupAxObject(b._objectname, "", b._r,"","");
           
        }

        void breadcrumb_l_MouseLeave(object sender, MouseEventArgs e)
        {
            Telerik.Windows.Controls.Label l = (Telerik.Windows.Controls.Label)sender;
            l.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#00FFFFFF")); ;
        }

        void breadcrumb_l_MouseEnter(object sender, MouseEventArgs e)
        {
            Telerik.Windows.Controls.Label l = (Telerik.Windows.Controls.Label)sender;
            l.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD1D1D1")); ;

        }

        void rB_Click(object sender, RoutedEventArgs e)
        {
            RadButton rB = (RadButton)sender;

            string tagdef = rB.Tag.ToString();
            string[] b = tagdef.Remove(0, "Custom*".Length).Split(':');

            string name = b[0];
            string type = (b.Length>1) ? b[1] : "";
            string action = (b.Length>2) ? b[2] : "";
            string recordtypes = (b.Length > 3) ? b[3] : "";
            string confirm = (b.Length > 4) ? b[4] : "";

            string objname = "";
            string id = "";
            string mattername = "";
            string message = "";
            
            if (type == "Data")
            {
                // get the id of the selected thing
                if (radGridView1.SelectedItem != null)
                {
                    DataRow r = ((DataRowView)radGridView1.SelectedItem).Row;
                    id = r["Id"].ToString();                    
                    objname = _sObjectDef.Name;
                }
                else
                {
                    message = "An item must be selected";
                }
                
            }
            else if (type == "Add")
            {
                // get the id of the parent thing
                id = _sObjectDef.ParentId;
                objname = _sObjectDef.Name;
                DataRow r = ((DataRowView)this._parentObject.radGridView1.SelectedItem).Row;
                mattername = r["Name"].ToString();
            }

            if (message != "")
            {
                MessageBox.Show(message);
            }
            else
            {
               
                if (confirm != "")
                {
                    MessageBoxResult result = MessageBox.Show(confirm, "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result != MessageBoxResult.Yes)
                    {
                        return;
                    }
                }

                // special actions ...
                if (action == "Template")
                {
                    // get the template name from the Matter - need to think about how to make this
                    // setting driven
                    DataRow r = ((DataRowView)_parentObject.radGridView1.SelectedItem).Row;
                    string templatename = "";
                    if (r.Table.Columns.Contains("Template__c"))
                    {
                        templatename = r["Template__c"].ToString();
                    }

                    this.CreateFromTemplate(objname, id, mattername, templatename);
                    return;
                }

                if (action == "ExistingTemplate")
                {
                    // get the template name from the Matter - need to think about how to make this
                    // setting driven
                    DataRow r = ((DataRowView)_parentObject.radGridView1.SelectedItem).Row;
                    string templatename = "";
                    if (r.Table.Columns.Contains("Template__c"))
                    {
                        templatename = r["Template__c"].ToString();
                    }

                    this.CreateFromExistingTemplate(objname, id, mattername, templatename);
                    return;
                }


                bsyInd.IsBusy = true;
                bsyInd.BusyContent = action;
                _buttonBackgroundWorker = new BackgroundWorker();
                _buttonBackgroundWorker.DoWork += (obj, ev) => buttonWorkerDoWork(obj, ev, action, objname, id);
                _buttonBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(buttonBackgroundWorker_RunWorkerCompleted);
                _buttonBackgroundWorker.RunWorkerAsync();

            }
            
        }



        void buttonBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _buttonBackgroundWorker.DoWork -= (obj, ev) => buttonWorkerDoWork(obj, ev, "", "", "");
            _buttonBackgroundWorker.RunWorkerCompleted -= buttonBackgroundWorker_RunWorkerCompleted;

            bsyInd.IsBusy = false;

            DataReturn dr = (DataReturn)e.Result;

            if (!dr.success)
            {
                MessageBox.Show(dr.strRtn);
            }
            else
            {

                // Refresh - this comes from the value returned from Salesforce
                // 1. if it "Object" - then just refresh the current object
                // 2. if it is "Parent" then refresh the parent object reselecting the current value
                //    this is for things like Task Accepts that update the Matter
                // 3. if it is "Filter-" then the name of the Filter then update the Parent *and* change 
                //    the filter on the list to the new value - this is for thing like clockstopper where
                //    a subobject gets created - have to be able to see it


                if (dr.reload.ToLower() == "object")
                {
                    this.LoadData(true);
                }
                else if (dr.reload.ToLower() == "parent")
                {
                    _selectidonceloaded = dr.id;

                    if (_parentObject != null)
                    {
                        _parentObject.LoadData(true);
                        this.LoadData(true);
                    }
                    else
                    {
                        this.LoadData(true);
                    }
                } 
                else if(dr.reload.ToLower().StartsWith("filter:")){

                    string selectfilter = dr.reload.Substring("Filter:".Length,dr.reload.Length-"Filter:".Length);
                    this.cbFilter.Text = selectfilter;
                    this._sObjectDef.SortColumn = "";
                    this._sObjectDef.CurrnetPage = 0;

                    _selectidonceloaded = dr.id;

                    if (_parentObject != null)
                    {
                        _parentObject.LoadData(true);
                        this.LoadData(true);
                    }
                    else
                    {
                        this.LoadData(true);
                    }
                }



            }

        }


        void buttonWorkerDoWork(object sender, DoWorkEventArgs e, string action,string objname,string id)
        {
            DataReturn dr = _d.Exec(action, objname, id);
            e.Result = dr;
        }


        private void NewButtonContent_ItemClick(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            RadMenuItem mi = e.OriginalSource as RadMenuItem;
            string RecordTypeName = mi.Header.ToString();
            string RecordTypeId = mi.Tag.ToString();

            RadContextMenu m = (RadContextMenu)(sender);
            m.IsOpen = false;

            MessageBoxResult res = MessageBox.Show("Do you want to add a new " + RecordTypeName + " item?", "Confirm", MessageBoxButton.OKCancel);
            if (res == MessageBoxResult.Cancel)
            {
                return;
            }

            //Add to the grid and select
            DataView v = (DataView)this.radGridView1.ItemsSource;
            DataRow newrow = v.Table.Rows.Add();
            foreach (DataRowView rw in (DataView)this.radGridView1.ItemsSource)
            {
                if (rw["Id"].ToString() == "") radGridView1.SelectedItem = rw;                       
            }

            // Populate the new row with the RecordType and any default values           
            string pfield = _sObjectDef.Parent;
            string pfieldid = _sObjectDef.Parent;
            if (!pfield.EndsWith("__c"))
            {
                pfieldid += "Id";
            }
            else
            {
                pfield = pfield.Replace("__c", "__r");
            }


            foreach (DataColumn dc in newrow.Table.Columns)
            {
                // SForceEdit.SObjectDef.FieldGridCol f = sObj.GetField(dc.ColumnName);
                if (dc.ColumnName == "RecordTypeId")
                {
                    newrow["RecordTypeId"] = RecordTypeId;
                }

                if (_sObjectDef.ParentId!=null && dc.ColumnName == pfieldid)
                {                    
                    newrow[dc.ColumnName] = _sObjectDef.ParentId;
                    newrow[pfield + "_Name"] = _sObjectDef.ParentName;
                    if (newrow.Table.Columns.Contains(pfield + "_Type")) newrow[pfield + "_Type"] = _parentObject._sObjectDef.Name;
                }

                // set the Owner to the current user
                if (dc.ColumnName == "OwnerId")
                {
                    newrow["OwnerId"] = _d.GetUserId();
                    newrow["Owner_Name"] = _d.GetUser();
                    if(newrow.Table.Columns.Contains("Owner_Type")) newrow["Owner_Type"] = "User";
                }
            }

            LoadRow(newrow);
        }


        private void CreateFromTemplate(string objname,string id,string name,string templatename)
        {
            //Contract axC = new Contract();
            //axC.CreateNewVersion(objname,id,templatename);

            NewFromTemplate axNfT = new NewFromTemplate();
            axNfT.Create(objname, id,name, templatename);
            axNfT.Show();

        }

        private void CreateFromExistingTemplate(string objname, string id, string name, string templatename)
        {
            Exsisting axNfT = new Exsisting();
            axNfT.Create(objname, id, name, templatename);
            axNfT.Show();
            axNfT.Focus();
        }


    }
}
