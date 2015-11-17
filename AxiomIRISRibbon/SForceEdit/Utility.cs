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
using System.Collections;
using System.Data;
using System.Collections.ObjectModel;

namespace AxiomIRISRibbon.SForceEdit
{
    class Utility
    {

        // Handy form functions - pass in a DataTable and the form and it will update any 
        // matching fields - to pass in the form pass all the Grid parents with form fields inside
        // did try to use the logical/visual trees but had lots of issues! visual tree only shows visible and logical didn't seem to work
        static public void UpdateForm(StackPanel sp, DataRow dr)
        {

            foreach (DependencyObject gboxes in sp.Children)
            {

                if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox) || gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander))
                {
                    //I'm sure there is a better way to do this! but allow the containter to be a Group Box or an Expander
                    UIElementCollection p1 = null;
                    if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox)) p1 = ((Grid)(((Telerik.Windows.Controls.GroupBox)gboxes).Content)).Children;
                    if (gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander)) p1 = ((Grid)(((Telerik.Windows.Controls.RadExpander)gboxes).Content)).Children;

                    foreach (DependencyObject child in p1)
                    {

                        if (child.GetType() == typeof(StackPanel))
                        {
                            StackPanel spcontrol = (StackPanel)child;
                            foreach (Object spchildControl in spcontrol.Children)
                            {
                                Object childControl = spchildControl;

                                if (childControl != null)
                                {
                                    //if there is a scroll then get the child element
                                    if (childControl.GetType() == typeof(ScrollViewer))
                                    {
                                        ScrollViewer sc = (ScrollViewer)childControl;
                                        if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox) || sc.Content.GetType() == typeof(RadAutoCompleteBox))
                                        {
                                            childControl = (Control)sc.Content;
                                        }
                                    }

                                    if (childControl.GetType() == typeof(TextBox))
                                    {
                                        TextBox tb = (TextBox)childControl;
                                        string updatable = "false";
                                        if (tb.Tag != null)
                                        {
                                            string[] n = tb.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());
                                            tb.Text = val;
                                            tb.ToolTip = (val == "" ? null : val);
                                        }

                                        // make readonly if row is null and readwrite if its not
                                        tb.IsReadOnly = !(dr == null ? false : (updatable == "True"));
                                    }
                                    else if (childControl.GetType() == typeof(ComboBox))
                                    {
                                        ComboBox cb = (ComboBox)childControl;
                                        string updatable = "false";
                                        if (cb.Tag != null)
                                        {
                                            string[] n = cb.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());
                                            cb.Text = val;
                                            cb.ToolTip = (val == "" ? null : val);
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        cb.IsReadOnly = !(dr == null ? false : (updatable == "True"));

                                    }
                                    else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadNumericUpDown))
                                    {
                                        RadNumericUpDown nb = (RadNumericUpDown)childControl;
                                        string updatable = "false";
                                        if (nb.Tag != null)
                                        {
                                            string[] n = nb.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());
                                            if (val != null && val != "")
                                            {
                                                nb.Value = Convert.ToDouble(val);
                                            }
                                            else
                                            {
                                                nb.Value = null;
                                            }
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        nb.IsEnabled = (dr == null ? false : (updatable == "True"));
                                    }
                                    else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadComboBox))
                                    {
                                        RadComboBox cb = (RadComboBox)childControl;
                                        string updatable = "false";
                                        if (cb.Tag != null)
                                        {
                                            string[] n = cb.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());

                                            if (name == "RecordTypeId")
                                            {
                                                foreach (RadComboBoxItem rbi in cb.Items)
                                                {
                                                    if (rbi.Tag.ToString() == val)
                                                    {
                                                        cb.SelectedItem = rbi;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                               
                                                // do what Salesforce does - if the value isn't in the list then jsut set it to the value anyway                                                                             
                                                // have to clear out any previous values so mark them with a tag

                                                // Clear out any temp values
                                                for (int z = cb.Items.Count-1; z >= 0; z--)
                                                {
                                                    
                                                    Object o = (Object)cb.Items[z];
                                                    if (o.GetType() == typeof(RadComboBoxItem))
                                                    {
                                                        RadComboBoxItem rbi = (RadComboBoxItem)o;
                                                        if (rbi.Tag != null)
                                                        {
                                                            if (rbi.Tag.ToString() == "TEMP")
                                                            {
                                                                cb.Items.RemoveAt(z);
                                                            }
                                                        }
                                                    }
                                                }


                                                if (!cb.IsEditable && val != "" && !cb.Items.Contains(val))
                                                {
                                                    RadComboBoxItem rbi = new RadComboBoxItem();
                                                    rbi.Content = val;
                                                    rbi.Tag = "TEMP";
                                                    cb.Items.Add(rbi);
                                                    cb.SelectedItem = rbi;

                                                }
                                                else
                                                {
                                                    // Intrestingly having a bit of an issue
                                                    // with setting the value of the combo with
                                                    // cb.Text = val - not working for Matter status
                                                    // step through and do it this way - think its cause
                                                    // I've got strings for the items rather than proper objects
                                                    
                                                    if (!cb.IsEditable)
                                                    {
                                                        cb.SelectedItem = null;
                                                        for (int z = 0; z < cb.Items.Count; z++)
                                                        {
                                                            if (cb.Items[z].GetType() == typeof(string))
                                                            {
                                                                if (cb.Items[z].ToString() == val)
                                                                {
                                                                    cb.SelectedIndex = z;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        cb.Text = val;
                                                    }
                                                    
                                                    // cb.Text = val;
                                                    cb.ToolTip = (val == "" ? null : val);

                                                }
                                            }
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        cb.IsEnabled = (dr == null ? false : (updatable == "True"));
                                    }
                                    else if (childControl.GetType() == typeof(AxSearchBox))
                                    {
                                        AxSearchBox sb = (AxSearchBox)childControl;
                                        if (sb.Tag != null)
                                        {
                                            sb.SetValue(dr);
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        sb.IsReadOnly(dr == null ? true : false);
                                    }
                                    else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadAutoCompleteBox))
                                    {
                                        RadAutoCompleteBox acb = (RadAutoCompleteBox)childControl;
                                        string updatable = "false";
                                        if (acb.Tag != null)
                                        {
                                            string[] n = acb.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());

                                            string[] vals = val.Split(';');

                                            ObservableCollection<string> cblist = (ObservableCollection<string>)acb.SelectedItems;
                                            cblist.Clear();
                                            acb.SearchText = "";

                                            foreach (string it in acb.ItemsSource)
                                            {
                                                if (Array.IndexOf(vals, it) > -1)
                                                {
                                                    cblist.Add(it);
                                                }
                                            }

                                            acb.ToolTip = (val == "" ? null : val);
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        acb.IsEnabled = (dr == null ? false : (updatable == "True"));
                                    }
                                    else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadDatePicker))
                                    {
                                        RadDatePicker dp = (RadDatePicker)childControl;
                                        string updatable = "false";
                                        if (dp.Tag != null)
                                        {
                                            string[] n = dp.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());
                                            if (val == "")
                                            {
                                                dp.SelectedValue = null;
                                            }
                                            else
                                            {
                                                dp.SelectedDate = Convert.ToDateTime(val);
                                            }
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        dp.IsEnabled = (dr == null ? false : (updatable == "True"));
                                    }
                                    else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadDateTimePicker))
                                    {
                                        RadDateTimePicker dp = (RadDateTimePicker)childControl;
                                        string updatable = "false";
                                        if (dp.Tag != null)
                                        {
                                            string[] n = dp.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "" : dr[name].ToString());
                                            if (val == "")
                                            {
                                                dp.SelectedValue = null;
                                            }
                                            else
                                            {
                                                dp.SelectedValue = Convert.ToDateTime(val);
                                            }
                                        }
                                        // make readonly if row is null and readwrite if its not
                                        dp.IsEnabled = (dr == null ? false : (updatable == "True"));
                                    }
                                    else if (childControl.GetType() == typeof(CheckBox))
                                    {
                                        CheckBox cbox = (CheckBox)childControl;
                                        string updatable = "false";
                                        if (cbox.Tag != null)
                                        {
                                            string[] n = cbox.Tag.ToString().Split('|');
                                            string name = n[0];
                                            updatable = n[1];
                                            string val = (dr == null ? "False" : dr[name].ToString());
                                            if (val == "") val = "false";
                                            cbox.IsChecked = Convert.ToBoolean(val);

                                        }
                                        // make readonly if row is null and readwrite if its not
                                        cbox.IsEnabled = (dr == null ? false : (updatable == "True"));
                                    }
                                }
                            }
                        }
                    }
                        
                }
            }

        }


        //Other way round - update the datarow from the form
        //returns if anything has changed, if nothing has we don't have to save the form
        static public bool UpdateRow(StackPanel sp, DataRow dr)
        {

            bool anychanges = false;
            if (sp != null)
            {
                foreach (DependencyObject gboxes in sp.Children)
                {

                    if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox) || gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander))
                    {
                        //I'm sure there is a better way to do this! but allow the containter to be a Group Box or an Expander
                        UIElementCollection p1 = null;
                        if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox)) p1 = ((Grid)(((Telerik.Windows.Controls.GroupBox)gboxes).Content)).Children;
                        if (gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander)) p1 = ((Grid)(((Telerik.Windows.Controls.RadExpander)gboxes).Content)).Children;


                        foreach (DependencyObject child in p1)
                        {


                            if (child.GetType() == typeof(StackPanel))
                            {
                                StackPanel spcontrol = (StackPanel)child;
                                foreach (Object spchildControl in spcontrol.Children)
                                {

                                    Object childControl = spchildControl;

                                    if (childControl != null)
                                    {
                                        //if there is a scroll then get the child element
                                        if (childControl.GetType() == typeof(ScrollViewer))
                                        {
                                            ScrollViewer sc = (ScrollViewer)childControl;
                                            if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox) || sc.Content.GetType() == typeof(RadAutoCompleteBox))
                                            {
                                                childControl = (Control)sc.Content;
                                            }
                                        }

                                        if (childControl.GetType() == typeof(TextBox))
                                        {
                                            TextBox tb = (TextBox)childControl;

                                            if (tb.Tag != null)
                                            {
                                                string name = tb.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    if (dr[name].ToString() != tb.Text)
                                                    {
                                                        dr[name] = tb.Text;
                                                        anychanges = true;
                                                    }
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(ComboBox))
                                        {
                                            ComboBox cb = (ComboBox)childControl;
                                            if (cb.Tag != null)
                                            {
                                                string name = cb.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    if (dr[name].ToString() != cb.Text)
                                                    {
                                                        dr[name] = cb.Text;
                                                        anychanges = true;
                                                    }
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(RadNumericUpDown))
                                        {
                                            RadNumericUpDown nb = (RadNumericUpDown)childControl;
                                            if (nb.Tag != null)
                                            {
                                                string name = nb.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    string val = nb.Value.ToString();
                                                    if (val == "")
                                                    {
                                                        if (dr[name] != DBNull.Value)
                                                        {
                                                            dr[name] = DBNull.Value;
                                                            anychanges = true;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (dr[name].ToString() != val)
                                                        {
                                                            dr[name] = val;
                                                            anychanges = true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadComboBox))
                                        {
                                            Telerik.Windows.Controls.RadComboBox cb = (Telerik.Windows.Controls.RadComboBox)childControl;
                                            if (cb.Tag != null)
                                            {
                                                string name = cb.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    //Currently got the RecordTypeId as a picklist with the id as the tag of the item
                                                    //may change later!
                                                    if (name == "RecordTypeId")
                                                    {
                                                        foreach (RadComboBoxItem rbi in cb.Items)
                                                        {
                                                            if (rbi.IsSelected)
                                                            {
                                                                if (dr[name].ToString() != rbi.Tag.ToString())
                                                                {
                                                                    dr[name] = rbi.Tag.ToString();
                                                                    anychanges = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (dr[name].ToString() != cb.Text)
                                                        {
                                                            dr[name] = cb.Text;
                                                            anychanges = true;
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                        else if (childControl.GetType() == typeof(AxSearchBox))
                                        {
                                            AxSearchBox sb = (AxSearchBox)childControl;
                                            //Update Value returns true if the value has been changed
                                            if (sb.UpdateValue(dr))
                                            {
                                                anychanges = true;
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadAutoCompleteBox))
                                        {
                                            RadAutoCompleteBox acb = (RadAutoCompleteBox)childControl;
                                            if (acb.Tag != null)
                                            {
                                                acb.SearchText = "";
                                                string name = acb.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    ObservableCollection<string> cblist = (ObservableCollection<string>)acb.SelectedItems;
                                                    string val = string.Join(";", cblist.ToArray<string>());
                                                    if (dr[name].ToString() != val)
                                                    {
                                                        dr[name] = val;
                                                        anychanges = true;
                                                    }
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadDatePicker))
                                        {
                                            RadDatePicker dp = (RadDatePicker)childControl;
                                            if (dp.Tag != null)
                                            {
                                                string name = dp.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    if (dp.SelectedDate != null)
                                                    {
                                                        if (dr[name] == DBNull.Value && dp.SelectedDate.Value != null)
                                                        {
                                                            dr[name] = dp.SelectedDate.Value;
                                                            anychanges = true;

                                                        }
                                                        else if (Convert.ToDateTime(dr[name]) != dp.SelectedDate.Value)
                                                        {
                                                            dr[name] = dp.SelectedDate.Value;
                                                            anychanges = true;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (dr[name] != DBNull.Value)
                                                        {
                                                            dr[name] = DBNull.Value;
                                                            anychanges = true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadDateTimePicker))
                                        {
                                            RadDateTimePicker dp = (RadDateTimePicker)childControl;
                                            if (dp.Tag != null)
                                            {

                                                string name = dp.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    if (dp.SelectedDate != null)
                                                    {
                                                        if (dr[name] == DBNull.Value && dp.SelectedDate.Value != null)
                                                        {
                                                            dr[name] = dp.SelectedDate.Value;
                                                            anychanges = true;

                                                        }
                                                        else if (Convert.ToDateTime(dr[name]) != dp.SelectedDate.Value)
                                                        {
                                                            dr[name] = dp.SelectedDate.Value;
                                                            anychanges = true;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (dr[name] != DBNull.Value)
                                                        {
                                                            dr[name] = DBNull.Value;
                                                            anychanges = true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(CheckBox))
                                        {
                                            CheckBox cb = (CheckBox)childControl;
                                            if (cb.Tag != null)
                                            {
                                                string name = cb.Tag.ToString().Split('|')[0];
                                                if (dr.Table.Columns.IndexOf(name) > -1)
                                                {
                                                    if (dr[name].ToString().ToLower() != cb.IsChecked.ToString().ToLower())
                                                    {
                                                        dr[name] = cb.IsChecked.ToString();
                                                        anychanges = true;
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
                return anychanges;
            
        }


        static public string CheckRequireFieldsHaveValues(StackPanel sp, SObjectDef s)
        {

            string message = "";
            if (sp != null)
            {
                foreach (DependencyObject gboxes in sp.Children)
                {

                    if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox) || gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander))
                    {
                        //I'm sure there is a better way to do this! but allow the containter to be a Group Box or an Expander
                        UIElementCollection p1 = null;
                        if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox)) p1 = ((Grid)(((Telerik.Windows.Controls.GroupBox)gboxes).Content)).Children;
                        if (gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander)) p1 = ((Grid)(((Telerik.Windows.Controls.RadExpander)gboxes).Content)).Children;


                        foreach (DependencyObject child in p1)
                        {


                            if (child.GetType() == typeof(StackPanel))
                            {
                                StackPanel spcontrol = (StackPanel)child;
                                foreach (Object spchildControl in spcontrol.Children)
                                {

                                    Object childControl = spchildControl;

                                    if (childControl != null)
                                    {
                                        //if there is a scroll then get the child element
                                        if (childControl.GetType() == typeof(ScrollViewer))
                                        {
                                            ScrollViewer sc = (ScrollViewer)childControl;
                                            if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox) || sc.Content.GetType() == typeof(RadAutoCompleteBox))
                                            {
                                                childControl = (Control)sc.Content;
                                            }
                                        }

                                        if (childControl.GetType() == typeof(TextBox))
                                        {
                                            TextBox tb = (TextBox)childControl;

                                            if (tb.Tag != null)
                                            {
                                                if (tb.Text == "")
                                                {
                                                    string name = tb.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(tb.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(ComboBox))
                                        {
                                            ComboBox cb = (ComboBox)childControl;
                                            if (cb.Tag != null)
                                            {
                                                if (cb.Text == "")
                                                {
                                                    string name = cb.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(cb.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(RadNumericUpDown))
                                        {
                                            RadNumericUpDown nb = (RadNumericUpDown)childControl;
                                            if (nb.Tag != null)
                                            {                                                
                                                if (nb.Value.ToString() == "")
                                                {
                                                    string name = nb.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(nb.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }                                                
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadComboBox))
                                        {
                                            Telerik.Windows.Controls.RadComboBox cb = (Telerik.Windows.Controls.RadComboBox)childControl;
                                            if (cb.Tag != null)
                                            {                                                
                                                if (cb.Text == "")
                                                {
                                                    string name = cb.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(cb.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }
                                            }

                                        }
                                        else if (childControl.GetType() == typeof(AxSearchBox))
                                        {
                                            AxSearchBox sb = (AxSearchBox)childControl;
                                            //Update Value returns true if the value has been changed
                                            string name = sb.Tag.ToString().Split('|')[0];
                                            bool required = Convert.ToBoolean(sb.Tag.ToString().Split('|')[2]);
                                            if (sb.Value == "")
                                            {
                                                if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                            }
                                            
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadAutoCompleteBox))
                                        {
                                            RadAutoCompleteBox acb = (RadAutoCompleteBox)childControl;
                                            if (acb.Tag != null)
                                            {
                                                acb.SearchText = "";
                                                
                                                ObservableCollection<string> cblist = (ObservableCollection<string>)acb.SelectedItems;
                                                string val = string.Join(";", cblist.ToArray<string>());
                                                if (val == "")
                                                {
                                                    string name = acb.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(acb.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadDatePicker))
                                        {
                                            RadDatePicker dp = (RadDatePicker)childControl;
                                            if (dp.Tag != null)
                                            {
                                                
                                                if (dp.SelectedDate == null)
                                                {
                                                    string name = dp.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(dp.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }

                                            }
                                        }
                                        else if (childControl.GetType() == typeof(Telerik.Windows.Controls.RadDateTimePicker))
                                        {
                                            RadDateTimePicker dp = (RadDateTimePicker)childControl;
                                            if (dp.Tag != null)
                                            {
                                                if (dp.SelectedDate == null)
                                                {
                                                    string name = dp.Tag.ToString().Split('|')[0];
                                                    bool required = Convert.ToBoolean(dp.Tag.ToString().Split('|')[2]);
                                                    if (required) message += (message == "" ? "" : ",") + s.GetField(name).Header;
                                                }
                                            }
                                        }
                                        else if (childControl.GetType() == typeof(CheckBox))
                                        {
                                            CheckBox cb = (CheckBox)childControl;
                                            if (cb.Tag != null)
                                            {
                                                string name = cb.Tag.ToString().Split('|')[0];
                                                // can't really be null!
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }


                }
            }
            return message;
        }


        // given the field name find the Search Box - this is so the label can get the values
        // when clicked on
        static public AxSearchBox FindAxSearch(StackPanel sp, string Name)
        {
            if (sp != null)
            {
                foreach (DependencyObject gboxes in sp.Children)
                {

                    if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox) || gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander))
                    {
                        //I'm sure there is a better way to do this! but allow the containter to be a Group Box or an Expander
                        UIElementCollection p1 = null;
                        if (gboxes.GetType() == typeof(Telerik.Windows.Controls.GroupBox)) p1 = ((Grid)(((Telerik.Windows.Controls.GroupBox)gboxes).Content)).Children;
                        if (gboxes.GetType() == typeof(Telerik.Windows.Controls.RadExpander)) p1 = ((Grid)(((Telerik.Windows.Controls.RadExpander)gboxes).Content)).Children;


                        foreach (DependencyObject child in p1)
                        {


                            if (child.GetType() == typeof(StackPanel))
                            {
                                StackPanel spcontrol = (StackPanel)child;
                                foreach (Object spchildControl in spcontrol.Children)
                                {

                                    Object childControl = spchildControl;

                                    if (childControl != null)
                                    {
                                        //if there is a scroll then get the child element
                                        if (childControl.GetType() == typeof(ScrollViewer))
                                        {
                                            ScrollViewer sc = (ScrollViewer)childControl;
                                            if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox) || sc.Content.GetType() == typeof(RadAutoCompleteBox))
                                            {
                                                childControl = (Control)sc.Content;
                                            }
                                        }


                                        if (childControl.GetType() == typeof(AxSearchBox))
                                        {
                                            AxSearchBox sb = (AxSearchBox)childControl;
                                            if (sb.GetIdFieldName() == Name) return sb;
                                        }
                                    }
                                }
                            }
                        }


                    }
                }
            }
            return null;
        }
    }






}
