using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using Telerik.Windows.Controls;
using System.Collections.ObjectModel;

namespace AxiomIRISRibbon.SForceEdit
{
    //Hold the stuff about the SObject that we need, mainly the FieldLists

    public class SObjectDef
    {
        public struct FieldGridCol
        {
            public string Name;
            public string QueryField;
            public string RelationshipName;
            public string Header;
            public string DataType;
            public bool Query;
            public bool Update;
            public bool Create;
            public bool InGrid;

            public sfPartner.Field SFField; // hold the full Sforce definition - need the others for the dummy fields like _Name or _Type that don't have sf definitions
          
            public bool DependantParent;
            public List<string> DependantFields;
            public string DependantController;
            public Dictionary<string, string> DependantList;
        }

        public struct FilterEntry
        {
            public string Name;
            public string SOQL;
            public bool Default;
            public string OrderBy;
        }

        string _name;
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        string _label;
        public string Label
        {
            get { return _label; }
            set { _label = value; }
        }
        string _plurallabel;
        public string PluralLabel
        {
            get { return _plurallabel; }
            set { _plurallabel = value; }
        }
        string _namefield;
        public string NameField
        {
            get { return _namefield; }
            set { _namefield = value; }
        }

        Dictionary<string, FieldGridCol> _fields;

        List<string> _gridcolumnfields;
        public List<string> GridColumnFields
        {
            get { return _gridcolumnfields; }
            set { _gridcolumnfields = value; }
        }

        // overide with the tabfilter values
        string _tabfilters;
        public string TabFilters
        {
            get { return _tabfilters; }
            set { _tabfilters = value; }
        }

        Dictionary<string, FilterEntry> _filters;
        public Dictionary<string, FilterEntry> GridFilters
        {
            get { return _filters; }
            set { _filters = value; }
        }

        public bool _justreload;
        public bool JustReload
        {
            get { return _justreload; }
            set { _justreload = value; }
        }

        //SelectedId - just to remember the row that was selected so we can reselect
        string _SelectedId;
        public string SelectedId
        {
            get { return _SelectedId; }
            set { _SelectedId = value; }
        }


        //Id - set this when we know the ID and just want the one row
        string _Id;
        public string Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        //Filter
        string _filter;
        public string Filter
        {
            get { return _filter; }
            set { _filter = value; }
        }

        //Search
        string _search;
        public string Search
        {
            get { return _search; }
            set { _search = value; }
        }

        //Paging stuff
        bool _paging;
        public bool Paging
        {
            get { return _paging; }
            set
            {
                _paging = value;
                if (value) AddPaging();
            }
        }

        int _totalrecords;
        public int TotalRecords
        {
            get { return _totalrecords; }
            set { _totalrecords = value; }
        }

        int _currentpage;
        public int CurrnetPage
        {
            get { return _currentpage; }
            set { _currentpage = value; }
        }

        int _recordsperpage;
        public int RecordsPerPage
        {
            get { return _recordsperpage; }
            set { _recordsperpage = value; }
        }

        //Sorting
        string _sortcolumn;
        public string SortColumn
        {
            get { return _sortcolumn; }
            set { _sortcolumn = value; }
        }
        string _sortqueryfield;
        public string SortQueryField
        {
            get { return _sortqueryfield; }
            set { _sortqueryfield = value; }
        }

        string _sortdir;
        public string SortDir
        {
            get { return _sortdir; }
            set { _sortdir = value; }
        }

        //Parent Value
        string _parent;
        public string Parent
        {
            get { return _parent; }
            set { _parent = value; }
        }

        string _parenttype;
        public string ParentType
        {
            get { return _parenttype; }
            set { _parenttype = value; }
        }

        string _parentid;
        public string ParentId
        {
            get { return _parentid; }
            set { _parentid = value; }
        }

        string _parentname;
        public string ParentName
        {
            get { return _parentname; }
            set { _parentname = value; }
        }

        string _url;
        public string Url
        {
            get { return _url; }
            set { _url = value; }
        }

        private Dictionary<string, string> _RecordTypes;
        public Dictionary<string, string> RecordTypes
        {
            get { return _RecordTypes; }
            set { _RecordTypes = value; }
        }

        private Dictionary<string, sfPartner.RecordTypeMapping> _RecordTypeMapping;
        public Dictionary<string, sfPartner.RecordTypeMapping> RecordTypeMapping
        {
            get { return _RecordTypeMapping; }
            set { _RecordTypeMapping = value; }
        }

        string _defaultRecordType;
        public string DefaultRecordType
        {
            get { return _defaultRecordType; }
            set { _defaultRecordType = value; }
        }

        string _defaultRecordTypeName;
        public string DefaultRecordTypeName
        {
            get { return _defaultRecordTypeName; }
            set { _defaultRecordTypeName = value; }
        }

        private Dictionary<string, StackPanel> _Layouts;
        public Dictionary<string, StackPanel> Layouts{
            get {return _Layouts; }
            set {_Layouts = value;}
        }

        private Dictionary<string, StackPanel> _SideBarLayouts;
        public Dictionary<string, StackPanel> SideBarLayouts
        {
            get { return _SideBarLayouts; }
            set { _SideBarLayouts = value; }
        }

        private Telerik.Windows.Controls.GridViewColumnCollection _GridColumns;
        public Telerik.Windows.Controls.GridViewColumnCollection Columns
        {
            get { return _GridColumns; }
            set { _GridColumns = value; }
        }

        private Action _FormFieldChanged;
        private Action<string> _SalesforceButtonHit;
        private Action<string> _OpenButtonHit;

        private System.Windows.Media.Color _gbborder;
        private bool _setgbborder = false;
        public void SetGBBorder(System.Windows.Media.Color c){
            _gbborder= c;
            _setgbborder = true;
        }

        private RadWindow _r;


        public SObjectDef(string name)
        {
            _name = name;
            _fields = new Dictionary<string, FieldGridCol>();

            _paging = false;
            _search = "";
            _parent = "";
            _filter = "";
            _Id = "";

            _Layouts = new Dictionary<string, StackPanel>();
            _SideBarLayouts = new Dictionary<string, StackPanel>();
            _RecordTypeMapping = new Dictionary<string, sfPartner.RecordTypeMapping>();
            _GridColumns = new Telerik.Windows.Controls.GridViewColumnCollection();
            _tabfilters = "";
            _filters = new Dictionary<string, FilterEntry>();
        }

        public SObjectDef(string name,RadWindow r)
        {
            _r = r;

            _name = name;
            _fields = new Dictionary<string, FieldGridCol>();

            _paging = false;
            _search = "";
            _parent = "";
            _filter = "";
            _Id = "";

            _Layouts = new Dictionary<string, StackPanel>();
            _SideBarLayouts = new Dictionary<string, StackPanel>();
            _RecordTypeMapping = new Dictionary<string, sfPartner.RecordTypeMapping>();
            _GridColumns = new Telerik.Windows.Controls.GridViewColumnCollection();
            _tabfilters = "";
            _filters = new Dictionary<string, FilterEntry>();
        }

        public void AddPaging()
        {
            _paging = true;
            _currentpage = 0;
            _totalrecords = 0;
            _recordsperpage = 100;
            _sortcolumn = "";
            _sortdir = "";
        }

        public void AddField(string name, string relation, string header, string datatype, bool query, bool update,bool create,sfPartner.Field sffield)
        {
            //name is passed with dot for refrences, replace with _ as thats what comes back from query
            string nodotname = name.Replace('.', '_');

            if (!_fields.ContainsKey(nodotname))
            {
                FieldGridCol f = new FieldGridCol();
                f.Name = nodotname;
                f.QueryField = name;
                f.RelationshipName = relation;
                f.Header = header;
                f.DataType = datatype;
                f.Query = query;
                f.Update = update;
                f.Create = create;
                f.SFField = sffield;
                _fields.Add(nodotname, f);
            }
        }

        public void UpdateDependantPickList(string name, string dependantController, Dictionary<string, string> dependantValues)
        {
            if (_fields.ContainsKey(name))
            {
                FieldGridCol f = _fields[name];
                f.DependantController = dependantController;
                f.DependantList = dependantValues;
                _fields.Remove(name);
                _fields[name] = f;
            }
        }

        public void AddParentDependant(string name, string dependantController)
        {
            if (_fields.ContainsKey(dependantController))
            {
                FieldGridCol f = _fields[dependantController];
                f.DependantParent = true;
                if (f.DependantFields == null) f.DependantFields = new List<string>();
                f.DependantFields.Add(name);
                _fields.Remove(dependantController);
                _fields[dependantController] = f;
            }
        }

        public bool FieldExists(string name)
        {
            return _fields.ContainsKey(name);
        }

        public FieldGridCol GetField(string name)
        {
            return _fields[name];
        }

        public string GetQueryList()
        {
            string querylist = "";
            foreach (FieldGridCol f in _fields.Values)
            {
                querylist += (querylist == "" ? "" : ",") + f.QueryField;
            }
            return querylist;
        }

        public string GetSearchClause()
        {
            string querylist = "(";
            foreach (string col in _gridcolumnfields)
            {
                FieldGridCol f = GetField(col);

                if (f.DataType == "string" || f.DataType == "picklist" || f.DataType == "combobox" || f.DataType == "textarea")
                {
                    querylist += (querylist == "(" ? "" : " OR ") + f.QueryField + " like '%" + AxiomIRISRibbon.Utility.FixUpSOQLString(_search) + "%'";
                }
                else if (f.DataType == "reference")
                {
                    querylist += (querylist == "(" ? "" : " OR ") + f.RelationshipName + ".Name" + " like '%" + AxiomIRISRibbon.Utility.FixUpSOQLString(_search) + "%'";
                }

            }
            return querylist + ")";
        }

        public DataTable CreateDataTable()
        {
            DataTable d = new DataTable();
            foreach (FieldGridCol f in _fields.Values)
            {
                System.Data.DataColumn c = new DataColumn(f.Name, GetDataColumnType(f.DataType));
                d.Columns.Add(c);
            }
            return d;
        }

        private Type GetDataColumnType(string DataType)
        {
            if (DataType == "date" || DataType == "datetime")
            {
                return typeof(DateTime);
            }
            else if (DataType == "double" || DataType == "currency")
            {
                return typeof(Double);
            }
            else
            {
                return typeof(String);
            }
        }

        //Step through the Layouts from Salesforce and build the 
        //stack panels that contain the layout
        public void BuildLayouts(Data d, Action FormFieldChanged)
        {
            _FormFieldChanged = FormFieldChanged;

            //get the sobject describe
            sfPartner.DescribeSObjectResult dsr = d.GetSObject(this.Name);

            this.Label = dsr.label;
            this.PluralLabel = dsr.labelPlural;

            //Find the name fields
            for (int x = 0; x < dsr.fields.Length; x++)
            {
                if (dsr.fields[x].nameField)
                {
                    this.NameField = dsr.fields[x].name;
                }
            }

            this.Url = dsr.urlDetail;

            // get the layout 
            sfPartner.DescribeLayoutResult dlr = d.GetLayout(this.Name);

            // Add the ID - always need the ID
            AddField("Id", "", "Id", "Id", true, false,false,null);

            // Get the Record Types - only the ones that have mappings are available
            _RecordTypes = new Dictionary<string, string>();

            if (dsr.recordTypeInfos != null)
            {
                foreach (sfPartner.RecordTypeInfo rti in dsr.recordTypeInfos)
                {
                    // do not include the Master record type - this always ends AAA - could be called something else
                    // if we were in a diferent language ( found this on stackexchange)
                    if (rti.available && !rti.recordTypeId.ToString().EndsWith("AAA"))
                    {
                        _RecordTypes.Add(rti.recordTypeId, rti.name);
                    }
                }
            }
            

            foreach (sfPartner.DescribeLayout layout in dlr.layouts)
            {
                StackPanel flds = new StackPanel();
                flds.Name = "Fields" + layout.id;

                // find the required fields - we layout from the detail layout but we want to mark the required fields
                // so need to go through the edit layout to find them
                List<string> requiredFields = new List<string>();
                sfPartner.DescribeLayoutSection[] editLayoutSectionList = layout.editLayoutSections;
                for (int z1 = 0; z1 < editLayoutSectionList.Length; z1++)
                {
                    for (int i1 = 0; i1 < editLayoutSectionList[z1].rows; i1++)
                    {
                        for (int j1 = 0; j1 < editLayoutSectionList[z1].columns; j1++)
                        {
                            sfPartner.DescribeLayoutItem li1 = editLayoutSectionList[z1].layoutRows[i1].layoutItems[j1];
                            if (li1.required)
                            {
                                for (int k1 = 0; k1 < li1.layoutComponents.Length; k1++)
                                {
                                    requiredFields.Add(li1.layoutComponents[k1].value);
                                }
                            }
                        }
                    }
                }


                sfPartner.DescribeLayoutSection[] detailLayoutSectionList = layout.detailLayoutSections;

                for (int z = 0; z < detailLayoutSectionList.Length; z++)
                {
                    sfPartner.DescribeLayoutSection ls = detailLayoutSectionList[z];
                    Telerik.Windows.Controls.GroupBox gb = new Telerik.Windows.Controls.GroupBox();
                    gb.Header = ls.heading;
                    if (_setgbborder) gb.BorderBrush = new System.Windows.Media.SolidColorBrush(_gbborder);
                    gb.Margin = new Thickness(3, 3, 3, 3);

                    //set up the grid
                    int rows = ls.rows;
                    int cols = ls.columns;

                    Grid g = new Grid();
                    for (int i = 0; i < rows; i++)
                    {
                        RowDefinition rd = new RowDefinition();
                        g.RowDefinitions.Add(rd);
                    }

                    for (int j = 0; j < cols; j++)
                    {
                        ColumnDefinition cd1 = new ColumnDefinition();
                        cd1.Width = new GridLength(125);
                        g.ColumnDefinitions.Add(cd1);
                        ColumnDefinition cd2 = new ColumnDefinition();
                        g.ColumnDefinitions.Add(cd2);
                    }


                    // add *all* the fields - change this from just adding the ones in the layouts
                    // problem was we had some in the grid that aren't on the layouts
                    sfPartner.Field f = null;
                    for (int x = 0; x < dsr.fields.Length; x++)
                    {
                        // if (dsr.fields[x].name == li.layoutComponents[layoutComponentIndex].value) f = dsr.fields[x];
                        f = dsr.fields[x];
                        if (f != null)
                        {

                            //1. Add Field Definition
                            AddField(f.name,
                                f.relationshipName,
                                f.label,
                                f.type.ToString(),
                                true,
                                f.updateable,
                                f.createable,
                                f
                                );

                            //If this is a Relation add extra field to the field list with the Name - can't update directly so set updateable to false
                            if (f.type == sfPartner.fieldType.reference)
                            {
                                if (f.relationshipName != null)
                                {
                                    AddField(f.relationshipName + ".Name",
                                        f.relationshipName,
                                        f.label,
                                        f.type.ToString(),
                                        true,
                                        false,
                                        false,
                                        null);


                                    //if the reference can apply to more than one type then add the type
                                    if (f.referenceTo.Length > 1)
                                    {
                                        AddField(f.relationshipName + ".Type",
                                            f.relationshipName,
                                            f.label,
                                            f.type.ToString(),
                                            true,
                                            false,
                                            false,
                                            null);
                                    }
                                }
                            }
                        }
                    }






                    //add the fields
                    for (int i = 0; i < rows; i++)
                    {
                        sfPartner.DescribeLayoutRow lr = ls.layoutRows[i];
                        for (int j = 0; j < cols; j++)
                        {
                            if (j < lr.layoutItems.Length)
                            {
                                sfPartner.DescribeLayoutItem li = lr.layoutItems[j];
                                if (li != null && li.layoutComponents != null)
                                {
                                    bool useLabelStack = false;
                                    bool required = false;
                                    StackPanel labelholder = new StackPanel();
                                    labelholder.Orientation = Orientation.Vertical;
                                    //labelholder.HorizontalAlignment = HorizontalAlignment.Stretch;                                    
                                    // labelholder.Width = 100;                                   
                                    
                                    StackPanel holder = new StackPanel();
                                    holder.Orientation = Orientation.Vertical;
                                    holder.HorizontalAlignment = HorizontalAlignment.Stretch;
                                    //loop through the layout components 
                                    for (int layoutComponentIndex = 0; layoutComponentIndex < li.layoutComponents.Length; layoutComponentIndex++)
                                    {

                                        //Get the field def
                                        f = null;
                                        for (int x = 0; x < dsr.fields.Length; x++)
                                        {
                                            if (dsr.fields[x].name == li.layoutComponents[layoutComponentIndex].value) f = dsr.fields[x];
                                        }
                                        
                                        if (li.layoutComponents.Length > 1)
                                        {
                                            useLabelStack = true;
                                        }

                                        if (f != null)
                                        {
                                            // if the field is in the required field list then update it
                                            if (requiredFields.Contains(f.name)) required = true;

                                            //2. Add Label
                                            labelholder.SetValue(Grid.RowProperty, i);
                                            labelholder.SetValue(Grid.ColumnProperty, (j * 2));

                                            if (f.type != sfPartner.fieldType.boolean)
                                            {

                                               
                                                Telerik.Windows.Controls.Label lbl = new Telerik.Windows.Controls.Label();

                                                TextBlock wrapblock = new TextBlock();
                                                wrapblock.TextWrapping = TextWrapping.WrapWithOverflow;
                                                wrapblock.Width = 110;
                                                wrapblock.Text = li.layoutComponents.Length == 0 ? li.label : f.label;                                                                                                
                                                wrapblock.HorizontalAlignment = HorizontalAlignment.Left;

                                                if (f.type == sfPartner.fieldType.reference && f.name != "RecordTypeId")
                                                {
                                                    wrapblock.Tag = f.name;
                                                    wrapblock.TextDecorations = System.Windows.TextDecorations.Underline;
                                                    wrapblock.Cursor = System.Windows.Input.Cursors.Hand;
                                                    // subtle :-)
                                                    // wrapblock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Blue);
                                                    wrapblock.MouseEnter += wrapblock_MouseEnter;
                                                    wrapblock.MouseLeave += wrapblock_MouseLeave;
                                                    wrapblock.MouseDown += wrapblock_MouseDown;
                                                }
                                               
                                                if (f.inlineHelpText != null && f.inlineHelpText != "")
                                                {
                                                    Grid helpgrid = new Grid();
                                                    ColumnDefinition helpcd1 = new ColumnDefinition();
                                                    helpcd1.Width = new GridLength(1, GridUnitType.Star);
                                                    helpgrid.ColumnDefinitions.Add(helpcd1);
                                                    ColumnDefinition helpcd2 = new ColumnDefinition();
                                                    helpcd2.Width = new GridLength(18);
                                                    helpgrid.ColumnDefinitions.Add(helpcd2);

                                                    Telerik.Windows.Controls.Label help = new Telerik.Windows.Controls.Label();
                                                    help.ToolTip = f.inlineHelpText;
                                                    help.Content = "?";
                                                    help.Margin = new Thickness(0, -4, 0, 0);
                                                    help.SetValue(Grid.ColumnProperty, 1);
                                                    wrapblock.SetValue(Grid.ColumnProperty, 0);

                                                    helpgrid.Children.Add(wrapblock);
                                                    helpgrid.Children.Add(help);

                                                    helpgrid.HorizontalAlignment = HorizontalAlignment.Stretch;
                                                    
                                                    lbl.Content = helpgrid;
                                                    
                                                }
                                                else
                                                {
                                                    lbl.Content = wrapblock;
                                                }

                                                lbl.VerticalAlignment = VerticalAlignment.Top;
                                                lbl.Margin = new Thickness(3, 3, 0, 0);
                                                lbl.SetValue(Grid.RowProperty, i);
                                                lbl.SetValue(Grid.ColumnProperty, (j * 2));

                                                

                                                if (useLabelStack)
                                                {
                                                    if (f.type == sfPartner.fieldType.textarea)
                                                    {
                                                        lbl.Height = 23 * 4;
                                                    }
                                                    else
                                                    {
                                                        lbl.Height = 23;
                                                    }
                                                    labelholder.Children.Add(lbl);

                                                }
                                                else
                                                {
                                                    g.Children.Add(lbl);

                                                }

                                            }

                                            

                                            //3. Add the field
                                            //if (addLayoutComponentToGrid)
                                            //{

                                            holder.SetValue(Grid.RowProperty, i);
                                            holder.SetValue(Grid.ColumnProperty, (j * 2) + 1);

                                            

                                                if (f.type == sfPartner.fieldType.reference)
                                                {
                                                    if (f.name == "RecordTypeId")
                                                    {
                                                        //Get the list of RecordTypes and set that as the data
                                                        //make it a pick list to test but set to readonly - will have to implement a special thing to change
                                                        RadComboBox cb = new RadComboBox();
                                                        cb.Height = 23;
                                                        cb.Margin = new Thickness(3, 3, 0, 0);
                                                        cb.Padding = new Thickness(8, -3, 0, 0);
                                                        cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                        cb.Tag = f.name+"|"+false+"|"+required;
                                                       
                                                        //cb.SetValue(Grid.RowProperty, i);
                                                        //cb.SetValue(Grid.ColumnProperty, (j * 2) + 1);

                                                        

                                                        foreach (sfPartner.RecordTypeInfo rti in dsr.recordTypeInfos)
                                                        {
                                                            if (rti.available && !rti.recordTypeId.ToString().EndsWith("AAA"))
                                                            {
                                                                RadComboBoxItem rbi = new RadComboBoxItem();
                                                                rbi.Content = rti.name;
                                                                rbi.Tag = rti.recordTypeId;
                                                                cb.Items.Add(rbi);
                                                            }
                                                        }

                                                        holder.Children.Add(cb);
                                                        cb.IsEnabled = false;
                                                       // cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                                                    }
                                                    else
                                                    {

                                                        SForceEdit.AxSearchBox ax = new SForceEdit.AxSearchBox(f);
                                                        //ax.SetValue(Grid.RowProperty, i);
                                                        //ax.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                        ax.Tag = f.relationshipName + "_Name" + "|" + f.updateable + "|" + required;
                                                        ax.SelectionChanged += new RoutedEventHandler(ax_SelectionChanged);
                                                        holder.Children.Add(ax);
                                                        
                                                        //StyleManager.SetTheme(cb, StyleManager.ApplicationTheme);
                                                        //StyleManager.SetTheme(referenceFind, StyleManager.ApplicationTheme);
                                                    }


                                                }
                                                else if (f.type == sfPartner.fieldType.picklist)
                                                {
                                                    string dependantField = "";
                                                    sfPartner.Field dependantF = null;
                                                    Dictionary<string, string> dependantValues = null;
                                                    if (f.dependentPicklist)
                                                    {
                                                        dependantField = f.controllerName;
                                                        dependantValues = new Dictionary<string, string>();
                                                        for (int x = 0; x < dsr.fields.Length; x++)
                                                        {
                                                            if (dsr.fields[x].name == dependantField) dependantF = dsr.fields[x];
                                                        }
                                                    }

                                                    RadComboBox cb = new RadComboBox();
                                                    cb.Height = 23;
                                                    cb.Margin = new Thickness(3, 3, 0, 0);
                                                    cb.Padding = new Thickness(8, -3, 0, 0);
                                                    cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    cb.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //cb.SetValue(Grid.RowProperty, i);
                                                    //cb.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    if (!f.updateable) cb.IsEnabled = false;

                                                    
                                                    
                                                    foreach (sfPartner.PicklistEntry ple in f.picklistValues)
                                                    {
                                                        cb.Items.Add(ple.value);

                                                        //If this is a dependant list then work out the values it is valid for
                                                        if (f.dependentPicklist)
                                                        {
                                                            string validfor = "";
                                                            byte[] b = ple.validFor;
                                                            if (dependantF.type == sfPartner.fieldType.picklist)
                                                            {
                                                                for (int k = 0; k < b.Length * 8; k++)
                                                                {
                                                                    if ((b[k >> 3] & (0x80 >> k % 8)) != (byte)0x00)
                                                                    {
                                                                        validfor += (validfor == "" ? "" : ";") + dependantF.picklistValues[k].value;
                                                                    }
                                                                }
                                                            }
                                                            else if (dependantF.type == sfPartner.fieldType.@boolean)
                                                            {
                                                                if ((b[1 >> 3] & (0x80 >> 1 % 8)) != (byte)0x00)
                                                                {
                                                                    validfor += (validfor == "" ? "" : ";") + true.ToString();
                                                                }
                                                                if ((b[0 >> 3] & (0x80 >> 0 % 8)) != (byte)0x00)
                                                                {
                                                                    validfor += (validfor == "" ? "" : ";") + false.ToString();
                                                                }
                                                            }
                                                            //Console.WriteLine(f.name + ">>" + ple.value + " Valid for " + dependantField + " :" + validfor);
                                                            dependantValues[ple.value] = validfor;
                                                        }
                                                    }

                                                    if (f.dependentPicklist)
                                                    {
                                                        UpdateDependantPickList(f.name, dependantField, dependantValues);
                                                        AddParentDependant(f.name, dependantField);
                                                    }


                                                    holder.Children.Add(cb);
                                                    cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                                                }
                                                else if (f.type == sfPartner.fieldType.combobox)
                                                {

                                                    RadComboBox cb = new RadComboBox();
                                                    cb.Height = 23;
                                                    cb.Margin = new Thickness(3, 3, 0, 0);
                                                    cb.Padding = new Thickness(8, -3, 0, 0);
                                                    cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    cb.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //cb.SetValue(Grid.RowProperty, i);
                                                    //cb.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    if (!f.updateable) cb.IsReadOnly = true;

                                                    //Combos you can type in whatever you like
                                                    cb.IsEditable = true;
                                                    cb.Padding = new Thickness(4, -2, 0, 0);
                                                  
                                                    foreach (sfPartner.PicklistEntry ple in f.picklistValues)
                                                    {
                                                        cb.Items.Add(ple.value);
                                                    }
                                                    
                                                    // add event for selection change
                                                    cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                                                    // and keyup - you can type in anything
                                                    cb.KeyUp += cb_KeyUp;
                                                    //cb.TextChanged += new TextChangedEventHandler(tb_TextChanged);

                                                    holder.Children.Add(cb);
                                                }
                                                else if (f.type == sfPartner.fieldType.multipicklist)
                                                {
                                                    ScrollViewer sc = new ScrollViewer();
                                                    sc.Margin = new Thickness(3, 3, 0, 0);
                                                    //sc.SetValue(Grid.RowProperty, i);
                                                    //sc.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    sc.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    sc.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
                                                    sc.VerticalScrollBarVisibility = ScrollBarVisibility.Hidden;


                                                    RadAutoCompleteBox acb = new RadAutoCompleteBox();
                                                    acb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    acb.VerticalAlignment = System.Windows.VerticalAlignment.Stretch;
                                                    acb.Margin = new Thickness(0, 0, 0, 0);

                                                    acb.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    acb.BorderThickness = new Thickness(0, 0, 0, 0);

                                                    if (!f.updateable) acb.IsEnabled = false;

                                                    acb.TextSearchMode = TextSearchMode.Contains;
                                                    acb.SelectionMode = Telerik.Windows.Controls.Primitives.AutoCompleteSelectionMode.Multiple;
                                                    acb.FilteringBehavior = new ShowAllFilteringBehavior();
                                                    ObservableCollection<string> cblist = new ObservableCollection<string>();
                                                    foreach (sfPartner.PicklistEntry ple in f.picklistValues)
                                                    {
                                                        cblist.Add(ple.value);
                                                    }
                                                    acb.ItemsSource = cblist;
                                                    acb.SelectedItems = new ObservableCollection<string>();
                                                    acb.SelectionChanged += new SelectionChangedEventHandler(acb_SelectionChanged);

                                                    sc.Content = acb;
                                                    holder.Children.Add(sc);

                                                    StyleManager.SetTheme(sc, StyleManager.ApplicationTheme);
                                                    acb.FontWeight = FontWeights.SemiBold;
                                                }
                                                else if (f.type == sfPartner.fieldType.boolean)
                                                {

                                                    CheckBox cb = new CheckBox();
                                                    cb.Height = 23;
                                                    cb.Margin = new Thickness(3, 3, 0, 0);
                                                    cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    cb.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //cb.SetValue(Grid.RowProperty, i);
                                                    //cb.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    cb.Content = f.label;
                                                    if (!f.updateable) cb.IsEnabled = false;
                                                    holder.Children.Add(cb);

                                                    cb.Checked += new RoutedEventHandler(cb_Checked);
                                                    cb.Unchecked += new RoutedEventHandler(cb_Unchecked);
                                                    StyleManager.SetTheme(cb, StyleManager.ApplicationTheme);
                                                }
                                                else if (f.type == sfPartner.fieldType.date)
                                                {

                                                    RadDatePicker dp = new RadDatePicker();
                                                    dp.Height = 23;
                                                    dp.Margin = new Thickness(3, 3, 0, 0);
                                                    dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    dp.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //dp.SetValue(Grid.RowProperty, i);
                                                    //dp.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    if (!f.updateable) dp.IsReadOnly = true;
                                                    holder.Children.Add(dp);
                                                    dp.SelectionChanged += new SelectionChangedEventHandler(dp_SelectionChanged);
                                                    dp.FontWeight = FontWeights.SemiBold;
                                                }
                                                else if (f.type == sfPartner.fieldType.datetime)
                                                {
                                                    RadDateTimePicker dp = new RadDateTimePicker();
                                                    dp.Height = 23;
                                                    dp.Margin = new Thickness(3, 3, 0, 0);
                                                    dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    dp.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //dp.SetValue(Grid.RowProperty, i);
                                                    //dp.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    if (!f.updateable) dp.IsReadOnly = true;
                                                    holder.Children.Add(dp);
                                                    dp.SelectionChanged += new SelectionChangedEventHandler(dp_SelectionChanged);
                                                    dp.FontWeight = FontWeights.SemiBold;
                                                }
                                                else if (f.type == sfPartner.fieldType.currency || f.type == sfPartner.fieldType.@double)
                                                {
                                                    RadNumericUpDown dp = new RadNumericUpDown();
                                                    dp.Height = 23;
                                                    dp.Margin = new Thickness(3, 3, 0, 0);
                                                    dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    dp.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    dp.ValueFormat = ValueFormat.Numeric;
                                                    dp.NumberDecimalDigits = f.scale;
                                                    //dp.SetValue(Grid.RowProperty, i);
                                                    //dp.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    if (!f.updateable) dp.IsEnabled = false;
                                                    holder.Children.Add(dp);
                                                    dp.ValueChanged += new EventHandler<RadRangeBaseValueChangedEventArgs>(dp_ValueChanged);
                                                    dp.FontWeight = FontWeights.SemiBold;

                                                }
                                                else if (f.type == sfPartner.fieldType.textarea)
                                                {

                                                    TextBox tb = new TextBox();
                                                    tb.VerticalContentAlignment = System.Windows.VerticalAlignment.Top;
                                                    tb.Height = 23 * 4;
                                                    tb.AcceptsReturn = true;
                                                    tb.Margin = new Thickness(3, 3, 0, 0);
                                                    tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    tb.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //tb.SetValue(Grid.RowProperty, i);
                                                    //tb.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    tb.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
                                                    tb.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
                                                    if (!f.updateable) tb.IsReadOnly = true;
                                                    holder.Children.Add(tb);
                                                    tb.TextChanged += new TextChangedEventHandler(tb_TextChanged);
                                                    StyleManager.SetTheme(tb, StyleManager.ApplicationTheme);
                                                    tb.FontWeight = FontWeights.SemiBold;
                                                }
                                                else
                                                {
                                                    if (f.type != sfPartner.fieldType.@string)
                                                    {
                                                        Console.WriteLine("Not doing-" + f.type.ToString());
                                                    }
                                                    TextBox tb = new TextBox();
                                                    tb.Height = 23;
                                                    tb.Margin = new Thickness(3, 3, 0, 0);
                                                   // tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    tb.HorizontalAlignment = HorizontalAlignment.Stretch;
                                                    tb.Tag = f.name + "|" + f.updateable + "|" + required;
                                                    //tb.SetValue(Grid.RowProperty, i);
                                                    //tb.SetValue(Grid.ColumnProperty, (j * 2) + 1);
                                                    if (!f.updateable) tb.IsReadOnly = true;
                                                    holder.Children.Add(tb);
                                                    tb.TextChanged += new TextChangedEventHandler(tb_TextChanged);
                                                    StyleManager.SetTheme(tb, StyleManager.ApplicationTheme);
                                                    tb.FontWeight = FontWeights.SemiBold;
                                                }
                                            }
                                        
                                    }

                                    if (useLabelStack && labelholder.Children.Count > 0) g.Children.Add(labelholder);

                                    // Add the required indicator
                                    if (required)
                                    {
                                        System.Windows.Shapes.Rectangle r = new System.Windows.Shapes.Rectangle();
                                        r.Width = 3;
                                        r.Height = 19;
                                        r.SetValue(Grid.RowProperty, i);
                                        r.SetValue(Grid.ColumnProperty, (j * 2) + 0);
                                        r.Margin = new Thickness(0,5, 0, 0);
                                        r.ToolTip = "Required Information";
                                        r.Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Red);
                                        r.HorizontalAlignment = HorizontalAlignment.Right;
                                        r.VerticalAlignment = VerticalAlignment.Top;
                                        g.Children.Add(r);
                                    }

                                    if( holder.Children.Count > 0) g.Children.Add(holder);
                                }
                            }
                        }
                    }


                    if (g.Children.Count > 0)
                    {
                        gb.Content = g;
                        flds.Children.Add(gb);
                    }


                }

                Layouts.Add(layout.id, flds);
            }



            //If the Layouts don't have the RecordTypeId then add it in
            if (!FieldExists("RecordTypeId"))
            {
                //Get the field def
                sfPartner.Field f = null;
                for (int x = 0; x < dsr.fields.Length; x++)
                {
                    if (dsr.fields[x].name == "RecordTypeId")
                    {
                        f = dsr.fields[x];
                        AddField(f.name,
                                                f.relationshipName,
                                                "RecordTypeId",
                                                f.type.ToString(),
                                                true,
                                                f.updateable,
                                                f.createable,
                                                f);
                    }
                }
            }


            //Add columns for the grid - for now just add those declared in settings - later could allow user to show
            //and then get that to save in settings ... or even better read from ListView - can only get this from MetaData API
            //though and need to be Admin - could build a batch?
            List<string> searchCols = new List<string>();
            string settingColumns = Globals.ThisAddIn.GetSettings(this.Name, "Columns");

            if (settingColumns == "")
            {
                settingColumns = "Name";
            }

            _gridcolumnfields = new List<string>();

            //Add the columns in order of the view
            foreach (string settingCol in settingColumns.Split('|'))
            {
                //Find the Field
                if (this.FieldExists(settingCol))
                {
                    //keep a list
                    _gridcolumnfields.Add(settingCol);

                    SForceEdit.SObjectDef.FieldGridCol fgc = this.GetField(settingCol);

                    GridViewDataColumn column = new GridViewDataColumn();

                    if (fgc.DataType == "reference")
                    {
                        column.DataMemberBinding = new System.Windows.Data.Binding(fgc.RelationshipName + "_Name");
                        column.Header = fgc.Header;
                        column.UniqueName = fgc.Name;  
                    }
                    else
                    {
                        column.DataMemberBinding = new System.Windows.Data.Binding(fgc.Name);
                        column.Header = fgc.Header;
                        column.UniqueName = fgc.Name;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
                    }
                        
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

                        string nformat = "0";
                        if (fgc.SFField != null)
                        {
                            if (fgc.SFField.scale > 0)
                            {
                                nformat = "0." + "".PadRight(fgc.SFField.scale, '0');
                            }
                        }

                        column.DataFormatString = nformat;
                        column.TextAlignment = TextAlignment.Right;
                    }
                    column.MaxWidth = 200;
                    _GridColumns.Add(column);
                }
            }

            //get the record layout mappings and set the default value
            if (dlr.recordTypeMappings != null)
            {
                foreach (sfPartner.RecordTypeMapping rtm in dlr.recordTypeMappings)
                {
                    RecordTypeMapping.Add(rtm.recordTypeId, rtm);

                    if (rtm.defaultRecordTypeMapping)
                    {
                        _defaultRecordType = rtm.layoutId;
                        _defaultRecordTypeName = rtm.name;
                    }
                }
            }


            // Add the filters, from the definition - unless there is a tab filter overide and if there is use that            
            _filters = new Dictionary<string, FilterEntry>();

            string settingFilters ="";
            if (_tabfilters == "")
            {
                settingFilters = Globals.ThisAddIn.GetSettings(this.Name, "Filters");
            }
            else
            {
                settingFilters = _tabfilters;
            }

            if (settingFilters == "")
            {
                // add the default My and All
                FilterEntry f = new FilterEntry();                
                f.Name = "My " + this._plurallabel;
                f.SOQL = "OwnerId = '{UserId}'";
                f.Default = true;
                f.OrderBy = "";
                _filters.Add(f.Name, f);

                f = new FilterEntry();
                f.Name = "All " + this._plurallabel;
                f.SOQL = "";
                f.Default = false;
                f.OrderBy = "";
                _filters.Add(f.Name, f);

            }
            else
            {
                foreach (string settingFilter in settingFilters.Split('|'))
                {
                    if (settingFilter != "")
                    {
                        string[] s = settingFilter.Split(':');
                        FilterEntry f = new FilterEntry();
                        f.Name = s[0];                        
                        f.SOQL = s[1];

                        // if the name is just My and the SOQL isn't defined then add in the default
                        // and change the name
                        if (f.Name == "My" && f.SOQL=="")
                        {
                            f.Name = "My " + this._plurallabel;
                            f.SOQL = "OwnerId = '{UserId}'";
                        }
                        if (f.Name == "All" && f.SOQL == "")
                        {
                            f.Name = "All " + this._plurallabel;
                            f.SOQL = "";
                        }


                        if (s[2].ToLower() == "yes" || s[2].ToLower() == "true")
                        {
                            f.Default = true;
                        }
                        else
                        {
                            f.Default = false;
                        }
                        f.OrderBy = s[3];

                        _filters.Add(f.Name, f);
                    }
                }
            }

        }

        void wrapblock_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {

            TextBlock t = (TextBlock)sender;

            // MessageBox.Show("Click!");
            Telerik.Windows.Controls.Label l = null;
            if(t.Parent.GetType()==typeof(System.Windows.Controls.Grid)){
                l = (Telerik.Windows.Controls.Label)((System.Windows.Controls.Grid)t.Parent).Parent;
            }
            else
            {
                l = (Telerik.Windows.Controls.Label)t.Parent;
            }

            Grid g1 = null;
            if (l.Parent.GetType() == typeof(System.Windows.Controls.StackPanel))
            {
                g1 = (Grid)((StackPanel)l.Parent).Parent;
            }
            else {
                g1 = (Grid)l.Parent;
            }
            
            System.Windows.Controls.GroupBox gb = (System.Windows.Controls.GroupBox)g1.Parent;
            StackPanel sp1 = (StackPanel)gb.Parent;

            string tag = t.Tag.ToString();
            AxSearchBox ab = Utility.FindAxSearch(sp1,tag);
            
            // Calling these popup windows with the details "Zoom"            
            if (ab.Id != "")
            {
                Globals.ThisAddIn.OpenZoomEditWindow(ab.SFType, ab.Id);
                e.Handled = true;
            }


        }

        void wrapblock_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            TextBlock t = (TextBlock)sender;
            t.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Black);
            e.Handled = true;
        }

        void wrapblock_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            TextBlock t = (TextBlock)sender;
            t.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Purple);
            e.Handled = true;
        }

        void cb_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            FieldChanged();           
        }

        public void AddColumns(RadGridView r1){
            foreach (Telerik.Windows.Controls.GridViewColumn c in _GridColumns)
            {
                r1.Columns.Add(c);
            }
        }

        void dp_ValueChanged(object sender, RadRangeBaseValueChangedEventArgs e)
        {
            FieldChanged();
        }

        void tb_TextChanged(object sender, TextChangedEventArgs e)
        {
            FieldChanged();
        }

        void dp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FieldChanged();
        }

        void acb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FieldChanged();
        }

        void ax_SelectionChanged(object sender, RoutedEventArgs e)
        {
            FieldChanged();
        }


        RadComboBox FindCombo(StackPanel CurrentSPFields,string name)
        {
            //For Regular its a stackpanel with multiple group boxes with grid below
            //for compact its a stackpanel with a single expando with grid below
            for (int i = 0; i < CurrentSPFields.Children.Count; i++)
            {

                if (CurrentSPFields.Children[i].GetType() == typeof(Telerik.Windows.Controls.GroupBox))
                {
                    Telerik.Windows.Controls.GroupBox gb = (Telerik.Windows.Controls.GroupBox)CurrentSPFields.Children[i];
                    Grid g = (Grid)gb.Content;

                    for (int j = 0; j < g.Children.Count; j++)
                    {
                        if (g.Children[j].GetType() == typeof(StackPanel))
                        {
                            StackPanel spcontrol = (StackPanel)g.Children[j];
                            foreach (Object spchildControl in spcontrol.Children)
                            {
                                Object childControl = spchildControl;
                                if (childControl.GetType() == typeof(RadComboBox))
                                    if (((RadComboBox)childControl).Tag.ToString().StartsWith(name + "|")) return (RadComboBox)childControl;

                            }
                        }

                    }
                }

                if (CurrentSPFields.Children[i].GetType() == typeof(Telerik.Windows.Controls.RadExpander))
                {
                    Telerik.Windows.Controls.RadExpander exp = (Telerik.Windows.Controls.RadExpander)CurrentSPFields.Children[i];
                    Grid g = (Grid)exp.Content;

                    for (int j = 0; j < g.Children.Count; j++)
                    {
                        if (g.Children[j].GetType() == typeof(StackPanel))
                        {
                            StackPanel spcontrol = (StackPanel)g.Children[j];
                            foreach (Object spchildControl in spcontrol.Children)
                            {
                                Object childControl = spchildControl;
                                if (childControl.GetType() == typeof(RadComboBox))
                                    if (((RadComboBox)childControl).Tag.ToString().StartsWith(name + "|")) return (RadComboBox)childControl;

                            }
                        }
                    }
                }
            }
            return null;
        }

        void cb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FieldChanged();

            //Work out dependency drop downs
            RadComboBox cb = (RadComboBox)sender;
            string name = cb.Tag.ToString().Split('|')[0];
            SForceEdit.SObjectDef.FieldGridCol f = GetField(name);

            //find the parent stack panel
            Grid g1 = (Grid)((StackPanel)cb.Parent).Parent;
            StackPanel sp1 = null;
            if (g1.Parent.GetType() == typeof(Telerik.Windows.Controls.GroupBox))
            {
                Telerik.Windows.Controls.GroupBox gb1 = (Telerik.Windows.Controls.GroupBox)g1.Parent;
                sp1 = (StackPanel)gb1.Parent;
            }
            else if (g1.Parent.GetType() == typeof(Telerik.Windows.Controls.RadExpander))
            {
                Telerik.Windows.Controls.RadExpander rb1 = (Telerik.Windows.Controls.RadExpander)g1.Parent;
                sp1 = (StackPanel)rb1.Parent;
            }

            if (f.DependantParent)
            {
                foreach (string child in f.DependantFields)
                {
                    SForceEdit.SObjectDef.FieldGridCol fChild = GetField(child);
                    RadComboBox childCombo = FindCombo(sp1, child);

                    //This is a bit of a long way round but just cause its held this way in SalesForce
                    //Clear the list and then step through the dependentlist and if the value of the parent matches one of the array entries then add it to the list
                    if (childCombo != null)
                    {
                        string oldVal = childCombo.Text;
                        childCombo.Items.Clear();
                        foreach (string item in fChild.DependantList.Keys)
                        {
                            string[] list = fChild.DependantList[item].Split(';');
                            if (e.AddedItems.Count > 0)
                            {
                                string val = "";
                                if (e.AddedItems[0].GetType() == typeof(string))
                                {
                                    val = e.AddedItems[0].ToString();
                                }
                                if (e.AddedItems[0].GetType() == typeof(RadComboBoxItem))
                                {
                                    val = ((RadComboBoxItem)e.AddedItems[0]).Content.ToString();
                                }

                                if (Array.IndexOf(list, val) >= 0) childCombo.Items.Add(item);
                            }
                        }
                        childCombo.Text = oldVal;
                    }
                }
            }
        }

        void cb_Unchecked(object sender, RoutedEventArgs e)
        {
            FieldChanged();

            //Work out dependency drop downs
            CheckBox cb = (CheckBox)sender;
            HandleCheckBoxDependency(cb, false);

        }

        void cb_Checked(object sender, RoutedEventArgs e)
        {
            FieldChanged();
            //Work out dependency drop downs
            CheckBox cb = (CheckBox)sender;
            HandleCheckBoxDependency(cb, true);
        }

        void HandleCheckBoxDependency(CheckBox cb, bool val)
        {
            //*TODO* haven't actually tested this - need to set up a test case in SForce
            //Work out dependency drop downs
            SForceEdit.SObjectDef.FieldGridCol f = this.GetField(cb.Tag.ToString().Split('|')[0]);

            //find the parent stack panel
            Grid g1 = (Grid)((StackPanel)cb.Parent).Parent;
            StackPanel sp1 = null;
            if (g1.Parent.GetType() == typeof(Telerik.Windows.Controls.GroupBox))
            {
                Telerik.Windows.Controls.GroupBox gb1 = (Telerik.Windows.Controls.GroupBox)g1.Parent;
                sp1 = (StackPanel)gb1.Parent;
            }
            else if (g1.Parent.GetType() == typeof(Telerik.Windows.Controls.RadExpander))
            {
                Telerik.Windows.Controls.RadExpander rb1 = (Telerik.Windows.Controls.RadExpander)g1.Parent;
                sp1 = (StackPanel)rb1.Parent;
            }

            if (f.DependantParent)
            {
                foreach (string child in f.DependantFields)
                {
                    SForceEdit.SObjectDef.FieldGridCol fChild = this.GetField(child);
                    RadComboBox childCombo = FindCombo(sp1,child);

                    //This is a bit of a long way round but just cause its held this way in SalesForce
                    //Clear the list and then step through the dependentlist and if the value of the parent matches one of the array entries then add it to the list
                    string oldVal = childCombo.Text;
                    childCombo.Items.Clear();
                    foreach (string item in fChild.DependantList.Keys)
                    {
                        string[] list = fChild.DependantList[item].Split(';');
                        if (Array.IndexOf(list, val.ToString()) >= 0) childCombo.Items.Add(item);

                    }
                    childCombo.Text = oldVal;
                }
            }
        }

        void FieldChanged()
        {
            //Call the delegate
            _FormFieldChanged();
        }

        //Step through the Layouts from Salesforce and build the StackPanel for the 
        //compact layout - this just contains the fields that match the settings
        public void BuildCompactLayouts(Data d, Action FormFieldChanged, Action<string> SalesforceButtonHit, Action<string> OpenButtonHit)
        {
            _FormFieldChanged = FormFieldChanged;
            _SalesforceButtonHit = SalesforceButtonHit;
            _OpenButtonHit = OpenButtonHit;

            //get the sobject describe
            sfPartner.DescribeSObjectResult dsr = d.GetSObject(this.Name);

            this.Label = dsr.label;
            this.PluralLabel = dsr.labelPlural;

            //Find the name fields
            for (int x = 0; x < dsr.fields.Length; x++)
            {
                if (dsr.fields[x].nameField)
                {
                    this.NameField = dsr.fields[x].name;
                }
            }

            this.Url = dsr.urlDetail;

            //get the layout 
            sfPartner.DescribeLayoutResult dlr = d.GetLayout(this.Name);

            //Add the ID - always need the ID
            AddField("Id", "", "Id", "Id", true, false,false,null);

            // add *all* the fields - change this from just adding the ones in the layouts
            // the formula lookup use the data stored here so get everything so they can 
            // reference it
            sfPartner.Field f = null;
            for (int x = 0; x < dsr.fields.Length; x++)
            {
                // if (dsr.fields[x].name == li.layoutComponents[layoutComponentIndex].value) f = dsr.fields[x];
                f = dsr.fields[x];
                if (f != null)
                {

                    //1. Add Field Definition
                    AddField(f.name,
                        f.relationshipName,
                        f.label,
                        f.type.ToString(),
                        true,
                        f.updateable,
                        f.createable,
                        f
                        );

                    //If this is a Relation add extra field to the field list with the Name - can't update directly so set updateable to false
                    if (f.type == sfPartner.fieldType.reference)
                    {
                        if (f.relationshipName != null)
                        {
                            AddField(f.relationshipName + ".Name",
                                f.relationshipName,
                                f.label,
                                f.type.ToString(),
                                true,
                                false,
                                false,
                                null);


                            //if the reference can apply to more than one type then add the type
                            if (f.referenceTo.Length > 1)
                            {
                                AddField(f.relationshipName + ".Type",
                                    f.relationshipName,
                                    f.label,
                                    f.type.ToString(),
                                    true,
                                    false,
                                    false,
                                    null);
                            }
                        }
                    }
                }
            }

            foreach (sfPartner.DescribeLayout layout in dlr.layouts)
            {
                StackPanel flds = new StackPanel();
                flds.Name = "Fields" + layout.id;
                flds.Tag = _name;

                sfPartner.DescribeLayoutSection[] detailLayoutSectionList = layout.detailLayoutSections;

                //All the matching fileds just go in one groupbox
                //could have an expander with multiple group boxes?

                //Compact - just one column - for now assume that they will all be included 
                //and use that for the rows - if they aren't on the recordtypes layout then they
                //won't be shown - might cause a scroll when we don't one - come back to!
                string compactFields = Globals.ThisAddIn.GetSettings(this.Name, "Compact");
                if (compactFields == "")
                {
                    compactFields = "Name";
                }

                List<string> compactFieldsList = compactFields.Split('|').ToList<string>();

               // Telerik.Windows.Controls.GroupBox gb = new Telerik.Windows.Controls.GroupBox();
               // gb.Header = _label;
               // if (_setgbborder) gb.BorderBrush = new System.Windows.Media.SolidColorBrush(_gbborder);
               // gb.Margin = new Thickness(3, 3, 3, 3);

                

                Telerik.Windows.Controls.RadExpander exp = new RadExpander();
                if (_setgbborder)
                {
                    exp.BorderBrush = new System.Windows.Media.SolidColorBrush(_gbborder);
                    exp.BorderThickness = new Thickness(2);
                }
                
                Telerik.Windows.Controls.Label lbl2 = new Telerik.Windows.Controls.Label();
                lbl2.Content = _label;
                
                    Grid g2 = new Grid();
                    Button sfbutton = new RadButton();
                    sfbutton.Margin = new Thickness(0, 0, 0, 0);
                    sfbutton.ToolTip = "Open In Salesforce";

                    Image icon = new Image();
                    icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/sf.ico", UriKind.Relative));
                    sfbutton.Content = icon;
                    sfbutton.Height = 22;
                    sfbutton.Width = 22;
                    sfbutton.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                    sfbutton.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                    sfbutton.Click += new RoutedEventHandler(sfbutton_Click);

                    Button openbutton = new RadButton();
                    openbutton.Margin = new Thickness(0, 0, 24, 0);
                    openbutton.ToolTip = "Open";

                    icon = new Image();
                    icon.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/open.png", UriKind.Relative));
                    openbutton.Content = icon;
                    openbutton.Height = 22;
                    openbutton.Width = 22;
                    openbutton.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                    openbutton.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                    openbutton.Click += new RoutedEventHandler(openbutton_Click);


                g2.Children.Add(lbl2);
                g2.Children.Add(sfbutton);
                g2.Children.Add(openbutton);

                exp.Header = g2;// _label;
                exp.Margin = new Thickness(3, 3, 3, 3);
                exp.Tag = _name;

                //set up the grid
                int rows = compactFieldsList.Count;
                if (compactFields == "All")
                {
                    //Need to work out how many rows we are going to have
                    rows=0;
                    for (int z = 0; z < detailLayoutSectionList.Length; z++)
                    {
                        sfPartner.DescribeLayoutSection ls = detailLayoutSectionList[z];
                        rows += (ls.rows * ls.columns);
                    }
                }
                int cols = 1;

                Grid g = new Grid();
                for (int i = 0; i < rows; i++)
                {
                    RowDefinition rd = new RowDefinition();
                    g.RowDefinitions.Add(rd);
                }

                for (int j = 0; j < cols; j++)
                {
                    ColumnDefinition cd1 = new ColumnDefinition();
                    cd1.Width = new GridLength(100);
                    g.ColumnDefinitions.Add(cd1);
                    ColumnDefinition cd2 = new ColumnDefinition();
                    g.ColumnDefinitions.Add(cd2);
                }

                int rowcount = 0;

                for (int z = 0; z < detailLayoutSectionList.Length; z++)
                {
                                       
                    sfPartner.DescribeLayoutSection ls = detailLayoutSectionList[z];
                    rows = ls.rows;
                    cols = ls.columns;
                   
                    //add the fields
                    for (int i = 0; i < rows; i++)
                    {
                        sfPartner.DescribeLayoutRow lr = ls.layoutRows[i];
                        for (int j = 0; j < cols; j++)
                        {
                            if (j < lr.layoutItems.Length)
                            {
                                sfPartner.DescribeLayoutItem li = lr.layoutItems[j];
                                if (li != null && li.layoutComponents != null)
                                {
                                    bool useLabelStack = false;
                                    StackPanel labelholder = new StackPanel();
                                    labelholder.Orientation = Orientation.Vertical;
                                    labelholder.HorizontalAlignment = HorizontalAlignment.Left;

                                    StackPanel holder = new StackPanel();
                                    holder.Orientation = Orientation.Vertical;
                                    holder.HorizontalAlignment = HorizontalAlignment.Stretch;

                                    //loop through the layout components 
                                    f = null;
                                    for (int layoutComponentIndex = 0; layoutComponentIndex < li.layoutComponents.Length; layoutComponentIndex++)
                                    {

                                        f = null;

                                        //only add the field to the layout if is in the compact field list or the field list is All
                                        if (compactFields=="All" || compactFieldsList.Contains(li.layoutComponents[layoutComponentIndex].value))
                                        {
                                            
                                            for (int x = 0; x < dsr.fields.Length; x++)
                                            {
                                                if (dsr.fields[x].name == li.layoutComponents[layoutComponentIndex].value) f = dsr.fields[x];
                                            }


                                            if (li.layoutComponents.Length > 1)
                                            {
                                                useLabelStack = true;
                                                //TODO
                                                //This is when you have extra stuff like CreatedBy/CreatedDate
                                                //the CreatedDate is the second component and I should add it to the grid somehow
                                                //for now just add to query so it can appear in the grid
                                                // addLayoutComponentToGrid = false;
                                            }
                                        }
                                        else
                                        {
                                            f = null; // set to null so we don't add it
                                        }
                                        
                                        if (f != null)
                                        {

                                            //2. Add Label
                                            if (f.type != sfPartner.fieldType.boolean)
                                            {
                                                labelholder.SetValue(Grid.RowProperty, rowcount);
                                                labelholder.SetValue(Grid.ColumnProperty, 0);

                                                Telerik.Windows.Controls.Label lbl = new Telerik.Windows.Controls.Label();

                                                TextBlock wrapblock = new TextBlock();
                                                wrapblock.TextWrapping = TextWrapping.WrapWithOverflow;
                                                wrapblock.Text = li.layoutComponents.Length == 0 ? li.label : f.label; 


                                                if (f.inlineHelpText != null && f.inlineHelpText != "")
                                                {
                                                    Grid helpgrid = new Grid();
                                                    ColumnDefinition helpcd1 = new ColumnDefinition();
                                                    helpcd1.Width = new GridLength(1, GridUnitType.Star);
                                                    helpgrid.ColumnDefinitions.Add(helpcd1);
                                                    ColumnDefinition helpcd2 = new ColumnDefinition();
                                                    helpcd2.Width = new GridLength(18);
                                                    helpgrid.ColumnDefinitions.Add(helpcd2);

                                                    Telerik.Windows.Controls.Label help = new Telerik.Windows.Controls.Label();
                                                    help.ToolTip = f.inlineHelpText;
                                                    help.Content = "?";
                                                    help.Margin = new Thickness(0, -4, 0, 0);
                                                    help.SetValue(Grid.ColumnProperty, 1);
                                                    wrapblock.SetValue(Grid.ColumnProperty, 0);

                                                    helpgrid.Children.Add(wrapblock);
                                                    helpgrid.Children.Add(help);

                                                    lbl.Content = helpgrid;
                                                }
                                                else
                                                {
                                                    lbl.Content = wrapblock;
                                                }

                                                lbl.VerticalAlignment = VerticalAlignment.Top;
                                                lbl.Margin = new Thickness(3, 3, 0, 0);
                                                lbl.SetValue(Grid.RowProperty, rowcount);
                                                lbl.SetValue(Grid.ColumnProperty, 0);

                                                if (useLabelStack)
                                                {
                                                    if (f.type == sfPartner.fieldType.textarea)
                                                    {
                                                        lbl.Height = 23 * 4;
                                                    }
                                                    else
                                                    {
                                                        lbl.Height = 23;
                                                    }
                                                    labelholder.Children.Add(lbl);
                                                }
                                                else
                                                {
                                                    g.Children.Add(lbl);
                                                }

                                            }


                                            // 3. Add the field
                                            // if (addLayoutComponentToGrid)
                                            // {
                                                holder.SetValue(Grid.RowProperty, rowcount);
                                                holder.SetValue(Grid.ColumnProperty, 1);

                                                if (f.type == sfPartner.fieldType.reference)
                                                {
                                                    if (f.name == "RecordTypeId")
                                                    {
                                                        //Get the list of RecordTypes and set that as the data
                                                        //make it a pick list to test but set to readonly - will have to implement a special thing to change
                                                        RadComboBox cb = new RadComboBox();
                                                        cb.Height = 23;
                                                        cb.Margin = new Thickness(3, 3, 0, 0);
                                                        cb.Padding = new Thickness(8, -3, 0, 0);
                                                        cb.Tag = f.name+"|false";
                                                        // cb.SetValue(Grid.RowProperty, rowcount);
                                                        // cb.SetValue(Grid.ColumnProperty, 1);
                                                        cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                        cb.IsEnabled = false;

                                                        foreach (sfPartner.RecordTypeInfo rti in dsr.recordTypeInfos)
                                                        {
                                                            RadComboBoxItem rbi = new RadComboBoxItem();
                                                            rbi.Content = rti.name;
                                                            rbi.Tag = rti.recordTypeId;
                                                            cb.Items.Add(rbi);
                                                        }

                                                        holder.Children.Add(cb);
                                                        cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);

                
                                                    }
                                                    else
                                                    {

                                                        SForceEdit.AxSearchBox ax = new SForceEdit.AxSearchBox(f);
                                                        //ax.SetValue(Grid.RowProperty, rowcount);
                                                        //ax.SetValue(Grid.ColumnProperty, 2);
                                                        ax.Tag = f.relationshipName + "_Name";
                                                        ax.SelectionChanged += new RoutedEventHandler(ax_SelectionChanged);
                                                        holder.Children.Add(ax);

                                                        //StyleManager.SetTheme(referenceFind, StyleManager.ApplicationTheme);
                                                    }


                                                }
                                                else if (f.type == sfPartner.fieldType.picklist)
                                                {
                                                    string dependantField = "";
                                                    sfPartner.Field dependantF = null;
                                                    Dictionary<string, string> dependantValues = null;
                                                    if (f.dependentPicklist)
                                                    {
                                                        dependantField = f.controllerName;
                                                        dependantValues = new Dictionary<string, string>();
                                                        for (int x = 0; x < dsr.fields.Length; x++)
                                                        {
                                                            if (dsr.fields[x].name == dependantField) dependantF = dsr.fields[x];
                                                        }
                                                    }

                                                    RadComboBox cb = new RadComboBox();
                                                    cb.Height = 23;
                                                    cb.Margin = new Thickness(3, 3, 0, 0);
                                                    cb.Padding = new Thickness(8, -3, 0, 0);
                                                    cb.Tag = f.name + "|" + f.updateable;
                                                    cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    //cb.SetValue(Grid.RowProperty, rowcount);
                                                    //cb.SetValue(Grid.ColumnProperty,  1);
                                                    if (!f.updateable) cb.IsEnabled = false;

                                                   
                                                    foreach (sfPartner.PicklistEntry ple in f.picklistValues)
                                                    {
                                                        cb.Items.Add(ple.value);

                                                        //If this is a dependant list then work out the values it is valid for
                                                        if (f.dependentPicklist)
                                                        {
                                                            string validfor = "";
                                                            byte[] b = ple.validFor;
                                                            if (dependantF.type == sfPartner.fieldType.picklist)
                                                            {
                                                                for (int k = 0; k < b.Length * 8; k++)
                                                                {
                                                                    if ((b[k >> 3] & (0x80 >> k % 8)) != (byte)0x00)
                                                                    {
                                                                        validfor += (validfor == "" ? "" : ";") + dependantF.picklistValues[k].value;
                                                                    }
                                                                }
                                                            }
                                                            else if (dependantF.type == sfPartner.fieldType.@boolean)
                                                            {
                                                                if ((b[1 >> 3] & (0x80 >> 1 % 8)) != (byte)0x00)
                                                                {
                                                                    validfor += (validfor == "" ? "" : ";") + true.ToString();
                                                                }
                                                                if ((b[0 >> 3] & (0x80 >> 0 % 8)) != (byte)0x00)
                                                                {
                                                                    validfor += (validfor == "" ? "" : ";") + false.ToString();
                                                                }
                                                            }
                                                            //Console.WriteLine(f.name + ">>" + ple.value + " Valid for " + dependantField + " :" + validfor);
                                                            dependantValues[ple.value] = validfor;
                                                        }
                                                    }

                                                    if (f.dependentPicklist)
                                                    {
                                                        UpdateDependantPickList(f.name, dependantField, dependantValues);
                                                        AddParentDependant(f.name, dependantField);
                                                    }


                                                    holder.Children.Add(cb);
                                                    cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);
                                                }
                                                else if (f.type == sfPartner.fieldType.combobox)
                                                {

                                                    RadComboBox cb = new RadComboBox();
                                                    cb.Height = 23;
                                                    cb.Margin = new Thickness(3, 3, 0, 0);
                                                    cb.Padding = new Thickness(8, -3, 0, 0);
                                                    cb.Tag = f.name + "|" + f.updateable;
                                                    cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    //cb.SetValue(Grid.RowProperty, rowcount);
                                                    //cb.SetValue(Grid.ColumnProperty, 1);
                                                    if (!f.updateable) cb.IsReadOnly = true;

                                                    //Combos you can type in whatever you like
                                                    cb.IsEditable = true;
                                                    cb.Padding = new Thickness(4, -2, 0, 0);

                                                    foreach (sfPartner.PicklistEntry ple in f.picklistValues)
                                                    {
                                                        cb.Items.Add(ple.value);
                                                    }
                                                    cb.SelectionChanged += new SelectionChangedEventHandler(cb_SelectionChanged);

                                                    holder.Children.Add(cb);
                                                }
                                                else if (f.type == sfPartner.fieldType.multipicklist)
                                                {
                                                    ScrollViewer sc = new ScrollViewer();
                                                    sc.Margin = new Thickness(3, 3, 0, 0);
                                                    //sc.SetValue(Grid.RowProperty, rowcount);
                                                    //sc.SetValue(Grid.ColumnProperty,1 );
                                                    sc.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    sc.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
                                                    sc.VerticalScrollBarVisibility = ScrollBarVisibility.Hidden;

                                                    RadAutoCompleteBox acb = new RadAutoCompleteBox();
                                                    acb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    acb.VerticalAlignment = System.Windows.VerticalAlignment.Stretch;
                                                    acb.Margin = new Thickness(0, 0, 0, 0);

                                                    acb.Tag = f.name + "|" + f.updateable;
                                                    acb.BorderThickness = new Thickness(0, 0, 0, 0);

                                                    if (!f.updateable) acb.IsEnabled = false;

                                                    acb.TextSearchMode = TextSearchMode.Contains;
                                                    acb.SelectionMode = Telerik.Windows.Controls.Primitives.AutoCompleteSelectionMode.Multiple;
                                                    acb.FilteringBehavior = new ShowAllFilteringBehavior();
                                                    ObservableCollection<string> cblist = new ObservableCollection<string>();
                                                    foreach (sfPartner.PicklistEntry ple in f.picklistValues)
                                                    {
                                                        cblist.Add(ple.value);
                                                    }
                                                    acb.ItemsSource = cblist;
                                                    acb.SelectedItems = new ObservableCollection<string>();
                                                    acb.SelectionChanged += new SelectionChangedEventHandler(acb_SelectionChanged);

                                                    sc.Content = acb;
                                                    holder.Children.Add(sc);

                                                    StyleManager.SetTheme(sc, StyleManager.ApplicationTheme);
                                                    acb.FontWeight = FontWeights.SemiBold;
                                                }
                                                else if (f.type == sfPartner.fieldType.boolean)
                                                {

                                                    CheckBox cb = new CheckBox();
                                                    cb.Height = 23;
                                                    cb.Margin = new Thickness(3, 3, 0, 0);
                                                    cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    cb.Tag = f.name + "|" + f.updateable;
                                                    //cb.SetValue(Grid.RowProperty, rowcount);
                                                    //cb.SetValue(Grid.ColumnProperty,1 );
                                                    cb.Content = f.label;
                                                    if (!f.updateable) cb.IsEnabled = false;
                                                    holder.Children.Add(cb);

                                                    cb.Checked += new RoutedEventHandler(cb_Checked);
                                                    cb.Unchecked += new RoutedEventHandler(cb_Unchecked);
                                                    StyleManager.SetTheme(cb, StyleManager.ApplicationTheme);
                                                }
                                                else if (f.type == sfPartner.fieldType.date)
                                                {

                                                    RadDatePicker dp = new RadDatePicker();
                                                    dp.Height = 23;
                                                    dp.Margin = new Thickness(3, 3, 0, 0);
                                                    dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    dp.Tag = f.name + "|" + f.updateable;
                                                    //dp.SetValue(Grid.RowProperty, rowcount);
                                                    //dp.SetValue(Grid.ColumnProperty, 1);
                                                    if (!f.updateable) dp.IsReadOnly = true;
                                                    holder.Children.Add(dp);
                                                    dp.SelectionChanged += new SelectionChangedEventHandler(dp_SelectionChanged);
                                                    dp.FontWeight = FontWeights.SemiBold;
                                                }
                                                else if (f.type == sfPartner.fieldType.datetime)
                                                {
                                                    RadDateTimePicker dp = new RadDateTimePicker();
                                                    dp.Height = 23;
                                                    dp.Margin = new Thickness(3, 3, 0, 0);
                                                    dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    dp.Tag = f.name + "|" + f.updateable;
                                                    //dp.SetValue(Grid.RowProperty, rowcount);
                                                    //dp.SetValue(Grid.ColumnProperty, 1);
                                                    if (!f.updateable) dp.IsReadOnly = true;
                                                    holder.Children.Add(dp);
                                                    dp.SelectionChanged += new SelectionChangedEventHandler(dp_SelectionChanged);
                                                    dp.FontWeight = FontWeights.SemiBold;
                                                }
                                                else if (f.type == sfPartner.fieldType.currency || f.type == sfPartner.fieldType.@double)
                                                {
                                                    RadNumericUpDown dp = new RadNumericUpDown();
                                                    dp.Height = 23;
                                                    dp.Margin = new Thickness(3, 3, 0, 0);
                                                    dp.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    dp.Tag = f.name + "|" + f.updateable;
                                                    dp.ValueFormat = ValueFormat.Numeric;
                                                    dp.NumberDecimalDigits = f.scale;
                                                    //dp.SetValue(Grid.RowProperty, rowcount);
                                                    //dp.SetValue(Grid.ColumnProperty, 1);
                                                    if (!f.updateable) dp.IsEnabled = false;
                                                    holder.Children.Add(dp);
                                                    dp.ValueChanged += new EventHandler<RadRangeBaseValueChangedEventArgs>(dp_ValueChanged);
                                                    dp.FontWeight = FontWeights.SemiBold;

                                                }
                                                else if (f.type == sfPartner.fieldType.textarea)
                                                {

                                                    TextBox tb = new TextBox();
                                                    tb.VerticalContentAlignment = System.Windows.VerticalAlignment.Top;
                                                    tb.Height = 23 * 4;
                                                    tb.AcceptsReturn = true;
                                                    tb.Margin = new Thickness(3, 3, 0, 0);
                                                    tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                    tb.Tag = f.name + "|" + f.updateable;
                                                    tb.SetValue(Grid.RowProperty, rowcount);
                                                    tb.SetValue(Grid.ColumnProperty, 1);
                                                    //tb.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
                                                    //tb.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
                                                    if (!f.updateable) tb.IsReadOnly = true;
                                                    holder.Children.Add(tb);
                                                    tb.TextChanged += new TextChangedEventHandler(tb_TextChanged);
                                                    StyleManager.SetTheme(tb, StyleManager.ApplicationTheme);
                                                    tb.FontWeight = FontWeights.SemiBold;
                                                }
                                                else
                                                {
                                                    if (f.type != sfPartner.fieldType.@string) Console.WriteLine("Not doing-" + f.type.ToString());
                                                    TextBox tb = new TextBox();
                                                    tb.Height = 23;
                                                    tb.Margin = new Thickness(3, 3, 0, 0);
                                                    tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                                                  //  tb.HorizontalAlignment = HorizontalAlignment.Stretch;
                                                    tb.Tag = f.name + "|" + f.updateable;
                                                   // tb.SetValue(Grid.RowProperty, rowcount);
                                                   // tb.SetValue(Grid.ColumnProperty,1 );
                                                    if (!f.updateable) tb.IsReadOnly = true;
                                                    holder.Children.Add(tb);
                                                    tb.TextChanged += new TextChangedEventHandler(tb_TextChanged);
                                                    StyleManager.SetTheme(tb, StyleManager.ApplicationTheme);
                                                    tb.FontWeight = FontWeights.SemiBold;
                                                }
                                                rowcount++;
                                            }
                                                                                        
                                        }
                                    if (useLabelStack && labelholder.Children.Count > 0) g.Children.Add(labelholder);
                                    if (holder.Children.Count > 0) g.Children.Add(holder);
                                }
                            }
                        }
                    }

                    
                }

                if (g.Children.Count > 0)
                {
                    exp.Content = g;
                    flds.Children.Add(exp);
                }
                
                _SideBarLayouts.Add(layout.id, flds);
            }



            //If the Layouts don't have the RecordTypeId then add it in
            if (!FieldExists("RecordTypeId"))
            {
                //Get the field def
                f = null;
                for (int x = 0; x < dsr.fields.Length; x++)
                {
                    if (dsr.fields[x].name == "RecordTypeId")
                    {
                        f = dsr.fields[x];
                        AddField(f.name,
                                                f.relationshipName,
                                                "RecordTypeId",
                                                f.type.ToString(),
                                                true,
                                                f.updateable,
                                                f.createable,
                                                f
                                                );
                    }
                }
            }



            //get the record layout mappings and set the default value
            foreach (sfPartner.RecordTypeMapping rtm in dlr.recordTypeMappings)
            {
                RecordTypeMapping.Add(rtm.recordTypeId, rtm);

                if (rtm.defaultRecordTypeMapping)
                {
                    _defaultRecordType = rtm.layoutId;

                }
            }


        }

        void openbutton_Click(object sender, RoutedEventArgs e)
        {
            //Fire the delegate if it defined
            if (_OpenButtonHit != null) _OpenButtonHit(_name);
        }

        void sfbutton_Click(object sender, RoutedEventArgs e)
        {
            //Fire the delegate if it defined
            if(_SalesforceButtonHit!=null) _SalesforceButtonHit(_name);
        }

        //Special Filter Behavior for multi-select - show all on no match
        public class ShowAllFilteringBehavior : FilteringBehavior
        {
            public override IEnumerable<object> FindMatchingItems(string searchText, System.Collections.IList items, IEnumerable<object> escapedItems, string textSearchPath, TextSearchMode textSearchMode)
            {
                var result = base.FindMatchingItems(searchText, items, escapedItems, textSearchPath, textSearchMode) as IEnumerable<object>;

                if (string.IsNullOrEmpty(searchText) || !result.Any())
                {
                    return ((IEnumerable<object>)items).Where(x => !escapedItems.Contains(x));
                }

                return result;
            }
        }
    }
}
