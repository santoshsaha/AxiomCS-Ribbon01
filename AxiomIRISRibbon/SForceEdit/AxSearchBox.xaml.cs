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
using Telerik.Windows.Controls;
using System.Data;

namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for AxSearchBox.xaml
    /// </summary>
    public partial class AxSearchBox : UserControl
    {
        sfPartner.Field _f;
        private Data _d;

        string _id;
        string _value;
        string _type;
        string _typeName;
        string _namefield;

        bool _gotdata;
        bool _triggerevents;

        public static readonly RoutedEvent ChangedEvent = EventManager.RegisterRoutedEvent("SelectionChanged", RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(AxSearchBox));
        public event RoutedEventHandler SelectionChanged
        {
            add { AddHandler(ChangedEvent, value); }
            remove { RemoveHandler(ChangedEvent, value); }
        }
        void RaiseChangeEvent()
        {
            RoutedEventArgs newEventArgs = new RoutedEventArgs(AxSearchBox.ChangedEvent);
            RaiseEvent(newEventArgs);
        }


        public AxSearchBox(sfPartner.Field f)
        {
            InitializeComponent();
            _f = f;
            _d = Globals.ThisAddIn.getData();

            // TODO for now assume that we are looking up by the Name field
            // this isn't always true, e.g. Task is Subject *but* the only way to do it
            // is to load the full definition of the object from Salesforce and step through all the 
            // fields and find the one with nameField set to true - I'm actually doing that for the 
            // object that we load but would need to look up the others when we get a reference
            _namefield = "Name";

            //If it can only be one object then hide the object picker
            if (f.referenceTo.Length == 1)
            {
                o1.Visibility = System.Windows.Visibility.Collapsed;
                coldefo1.Width = new GridLength(0);                
            }

            //Otherwise populate
            o1.Items.Clear();
            foreach (string s in f.referenceTo)
            {
                RadComboBoxItem cbi = new RadComboBoxItem();
                cbi.Tag = s;
                cbi.Content = _d.GetSObjectDef(s).label;

                // hard code overide for the Group versus Queue
                // can't work out how this is filtered and where you can find it in the object definition or layout
                // definition just says "Group" but UI says Queue and shows Groups where Type = Queue - odd!
                // also the Type is returned as Queue from the SOQL
                if (s == "Group")
                {
                    cbi.Tag = "Queue";
                    cbi.Content = "Queue";
                }


                o1.Items.Add(cbi);
            }

            //set to the first one *might be a default - should use that
            _type = f.referenceTo[0];
            _typeName = _d.GetSObjectDef(_type).label;

            //Set special filter
            acb1.FilteringBehavior = new ShowAllFilteringBehavior();

            _gotdata = false;
            _triggerevents = true;

        }

        public void IsReadOnly(bool v){
            if (_f.updateable)
            {
                acb1.IsEnabled = !v;
                b1.IsEnabled = !v;
                o1.IsEnabled = !v;
            }
            else
            {
                acb1.IsEnabled = false;
                b1.IsEnabled = false;
                o1.IsEnabled = false;
            }
        }

        private void tbSearchButton_Click(object sender, RoutedEventArgs e)
        {
            acb1.Focus();
            //acb1.SearchText = "";
            GetData();
            acb1.Populate("");
        }

        public string SelectedItem
        {
            get { return acb1.SelectedItem.ToString(); }
            set {
                acb1.SelectedItem = null;
                acb1.SearchText = value; 
            }
        }

        public void SetValue(string id, string value, string type)
        {
            _id = id;
            _value = value;
            _type = type;
            this.SelectedItem = value;
        }

        public void SetValue(DataRow r)
        {

            if (r != null)
            {
                _triggerevents = false;

                string relationshipName = _f.relationshipName;
                string idname = relationshipName + "Id";
                if (relationshipName.EndsWith("__r"))
                {
                    idname = relationshipName.Substring(0, relationshipName.Length - 3) + "__c";
                }
                string valname = relationshipName + "_Name";
                string typename = relationshipName + "_Type";

                if (_f.referenceTo.Length > 1)
                {

                    // if the type has changed we have to get the data
                    if (_type != r[typename].ToString())
                    {
                        _gotdata = false;
                    }

                    _type = r[typename].ToString();
                    foreach (RadComboBoxItem i in o1.Items)
                    {
                        if (i.Tag.ToString() == _type)
                        {
                            _typeName = i.Content.ToString();
                            this.o1.SelectedItem = i;
                        }
                    }
                }

                _id = r[idname].ToString();
                _value = r[valname].ToString();

                // acb1.SelectedItem = null;
                acb1.SearchText = _value;
                this.ToolTip = (_value == "" ? null : _value);

                _triggerevents = true;
            }
            else
            {
                _triggerevents = false;
                _id = "";
                _value = "";
                _type = "";
                this.acb1.SearchText = _value;
                this.o1.SelectedItem = null;
                this.ToolTip = (_value == "" ? null : _value);
                _triggerevents = true;

            }
            }

        public bool UpdateValue(DataRow r)
        {
            bool changed = false;

            string relationshipName = _f.relationshipName;
            string idname = relationshipName + "Id";
            if (relationshipName.EndsWith("__r"))
            {
                idname = relationshipName.Substring(0, relationshipName.Length - 3) + "__c";
            }
            string valname = relationshipName + "_Name";
            string typename = relationshipName + "_Type";

            if (r[idname].ToString() != _id)
            {
                r[idname] = _id;
                r[valname] = _value;
                if (r.Table.Columns.Contains(typename)) r[typename] = _type;
                changed = true;
            }
                        
            return changed;
        }

        public string GetIdFieldName()
        {
            string relationshipName = _f.relationshipName;
            string idname = relationshipName + "Id";
            if (relationshipName.EndsWith("__r"))
            {
                idname = relationshipName.Substring(0, relationshipName.Length - 3) + "__c";
            }
            return idname;
        }

        public string Id
        {
            get { return _id; }
        }
        public string Value
        {
            get { return _value; }
        }
        public string SFType
        {
            get { return _type; }
        }


        public new double Width
        {
            get { return g1.Width; }
            set
            {
                //Not totally sure why but this works!
                g1.Width = value + 4;
                sp1.Width = Width;
                acb1.Width = Width;
            }

        }

        public new string ToolTip
        {
            get { return (string)g1.ToolTip; }
            set
            {
                g1.ToolTip = "Reference:" + _type + " Value:" + value;
            }
        }

        //This isn't right but it'll do for now - once I've got the soruce to the telerik copy what they do?
        private void acb1_MouseEnter(object sender, MouseEventArgs e)
        {
            if (acb1.IsFocused)
            {
                b1.BorderBrush = SystemColors.ControlDarkDarkBrush;
            }
            else
            {
                b1.BorderBrush = SystemColors.ControlDarkBrush;
            }
        }

        private void acb1_MouseLeave(object sender, MouseEventArgs e)
        {
            if (acb1.IsFocused)
            {
                b1.BorderBrush = SystemColors.ControlDarkBrush;
            }
            else
            {
                b1.BorderBrush = acb1.BorderBrush;
            }
        }

        private void acb1_GotFocus(object sender, RoutedEventArgs e)
        {
            //if (b1.BorderBrush != SystemColors.ControlDarkDarkBrush) b1.BorderBrush = SystemColors.ControlBrush;
            b1.BorderBrush = SystemColors.ControlDarkBrush;
            (acb1.ChildrenOfType<TextBox>().First() as TextBox).SelectAll();                       
        }

        private void acb1_LostFocus(object sender, RoutedEventArgs e)
        {
            b1.BorderBrush = acb1.BorderBrush;

            //Check if we have a matching value
            if (acb1.SearchText == "")
            {
                _value = "";
                _id = "";
            }
            else if (acb1.SelectedItem == null)
            {
                acb1.SearchText = _value;
            }

        }

        private void acb1_SearchTextChanged(object sender, EventArgs e)
        {
            //Check if we don't have a dataset get one!
            if (_triggerevents)
            {
                GetData();
                RaiseChangeEvent();
            }

        }

        //Get the data if we haven't already got it
        private void GetData()
        {
            if (!_gotdata)
            {

                // hard code for type = "Queue" - this is actually Group filter for Type = Queue
                // can't work out how to handle this from the API so just hardcoding
                if (_type == "Queue")
                {
                    SObjectDef sd = new SObjectDef("Group");
                    sd.AddField("Id", "", "Id", "Id", true, false, false,null);
                    sd.AddField(_namefield, "", "Name", "@string", true, false, false, null);
                    sd.AddField("Type", "", "Name", "@string", true, false, false, null);

                    DataTable dt = _d.GetData(sd).dt;
                    dt.DefaultView.Sort = _namefield;
                    dt.DefaultView.RowFilter = _namefield + "<>'' and type = 'Queue'";

                    acb1.ItemsSource = dt.DefaultView;
                    acb1.TextSearchPath = _namefield;
                    _gotdata = true;
                }
                else
                {

                    SObjectDef sd = new SObjectDef(_type);
                    sd.AddField("Id", "", "Id", "Id", true, false, false, null);
                    sd.AddField(_namefield, "", "Name", "@string", true, false, false, null);

                    DataTable dt = _d.GetData(sd).dt;
                    dt.DefaultView.Sort = _namefield;
                    dt.DefaultView.RowFilter = _namefield + "<>''";

                    acb1.ItemsSource = dt.DefaultView;
                    acb1.TextSearchPath = _namefield;
                    _gotdata = true;
                }
                
            }
        }


        private void o1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_triggerevents)
            {
                if (o1.SelectedItem != null)
                {
                    RadComboBoxItem cbi = ((RadComboBoxItem)o1.SelectedItem);
                    _type = cbi.Tag.ToString();
                    _typeName = cbi.Content.ToString();
                }
                else
                {

                }

                //wipe the value
                _value = "";
                _id = "";
                SelectedItem = "";
                _gotdata = false;
                RaiseChangeEvent();
            }
        }

        private void acb1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_triggerevents)
            {
                DataRowView r = (DataRowView)acb1.SelectedItem;
                if (r != null)
                {
                    _id = r["Id"].ToString();
                    _value = r["Name"].ToString();
                }
                else
                {
                    _id = "";
                    _value = "";
                }
                acb1.SearchText = _value;
                RaiseChangeEvent();
            }
        }


        private void acb1_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!acb1.IsKeyboardFocusWithin)
            {
                (acb1.ChildrenOfType<TextBox>().First() as TextBox).SelectAll();   
                e.Handled = true;
                acb1.Focus();
            }
        }

        //Special Filter Behavior show all if blank
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
