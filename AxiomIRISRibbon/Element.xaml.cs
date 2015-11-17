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
using Office = Microsoft.Office.Core;


namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for Element.xaml
    /// 
    /// </summary>
    public partial class Element : Window
    {
        private Data _d;
        string _clauseid;
        string _clausename;

        public Element()
        {
            InitializeComponent();
            Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();

            //Get the drop down values and populate the list box
            
            RefreshElementList();
            GetDropDown();


            if (!Globals.ThisAddIn.getDebug()) tbHidden.Visibility = System.Windows.Visibility.Hidden;


            btnCancel.IsEnabled = false;
            btnSave.IsEnabled = false;
            btnDelete.IsEnabled = false;
        }

        public void Refresh()
        {
            this.RefreshElementList();
        }


        private void RefreshElementList()
        {
            DataReturn dr = Utility.HandleData(_d.GetElements());
            if (!dr.success) return;
            dgElements.ItemsSource = dr.dt.DefaultView;
        }

        private void GetDropDown()
        {
            DataReturn dr = Utility.HandleData(_d.GetPickListValues("RibbonElement__c", "Type__c", false));
            if (!dr.success) return;
            DataTable dt = dr.dt;

            cbType.Items.Clear();
            ComboBoxItem i = new ComboBoxItem();
            i.Content = "";
            cbType.Items.Add(i);

            foreach (DataRow r in dt.Rows)
            {
                i = new ComboBoxItem();
                i.Content = r["Value"].ToString();
                cbType.Items.Add(i);
            }
        }

        private void btnReload_Click(object sender, RoutedEventArgs e)
        {
            RefreshElementList();
        }

        private void dgElements_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgElements.SelectedIndex > -1)
            {
                if (btnSave.IsEnabled)
                {
                    MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.Cancel)
                    {
                        dgElements.SelectedIndex = -1;
                        return;
                    }
                }

                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2, formGrid3 }, ((DataRowView)dgElements.SelectedItem).Row);
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
                btnDelete.IsEnabled = true;
            }
            else
            {
                btnDelete.IsEnabled = false;
                Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
                btnSave.IsEnabled = false;
                btnCancel.IsEnabled = false;
            
            }
        }


        private void ElementRowDoubleClick(object sender, RoutedEventArgs e)
        {
            this.Hide();
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


        public void NewElement(string Name, string Desc, string Text, string Xml, String ClauseId, String ClauseName)
        {
            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3});
            dgElements.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();

            //set the form values
            tbName.Text = Name;
            tbDescription.Text = Desc;

            //remember the templateid that we were called from
            _clauseid = ClauseId;
            _clausename = ClauseName;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            Save();

            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;

            RefreshElementList();
        }


        private string Save()
        {
            //check the required fields
            if (tbName.Text == "")
            {
                MessageBox.Show("Name field is required", "Problem", MessageBoxButton.OK);
                return "";
            }
            if (cbType.Text == "")
            {
                MessageBox.Show("Type field is required", "Problem", MessageBoxButton.OK);
                return "";
            }

            //Create a row from the list box to save the values
            DataRow drow;
            DataView dv = (DataView)dgElements.ItemsSource;
            drow = dv.Table.NewRow();
            //Update from the form
            Utility.UpdateRow(new Grid[] { formGrid1, formGrid2, formGrid3 }, drow);

            //Save the values         
            DataReturn dr = Utility.HandleData(_d.SaveElement(drow));
            string id = dr.id;
            tbId.Text = id;

            return id;
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
                Utility.UpdateForm(new Grid[] { formGrid1, formGrid2, formGrid3 }, ((DataRowView)dgElements.SelectedItem).Row);
            }
            btnSave.IsEnabled = false;
            btnCancel.IsEnabled = false;
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnInsert_Click(object sender, RoutedEventArgs e)
        {
            if (Globals.ThisAddIn.isClause())
            {

                Globals.ThisAddIn.ProcessingStart("Insert Element");

                //save - exit if we don't get an id back
                string elementid = tbId.Text;
                if (btnSave.IsEnabled)
                {
                    Globals.ThisAddIn.ProcessingUpdate("Save Element");
                    elementid =Save();
                    if (elementid == "")
                    {
                        return;
                    }
                }

                //Add Element to Clause
                string clauseid = Globals.ThisAddIn.GetCurrentDocId();
                //Check if the element already exists in the clause
                DataReturn dr = Utility.HandleData(_d.GetElement(clauseid, elementid));
                if (!dr.success) return;

                if (dr.dt.Rows.Count == 0)
                {
                    Globals.ThisAddIn.ProcessingUpdate("Add Element To Clause");                    
                    string name = Utility.Truncate(_clausename,35) + "-" + Utility.Truncate(tbName.Text,35);
                    dr = Utility.HandleData(_d.SaveClauseElement("", name, clauseid, elementid, "0"));
                    if (!dr.success) return;
                    string id = dr.id;
                }

                //Insert in the doc
                Globals.ThisAddIn.ProcessingUpdate("Insert Into Clause");
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Selection sel = Globals.ThisAddIn.Application.Selection;

                //Check we have the Text Style
                try
                {
                    if (!Utility.StyleExists(doc.Styles,"ContentControl"))
                    {
                        Word.Style s = doc.Styles.Add("ContentControl");
                        s.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightOrange;
                    }
                }
                catch (Exception)
                {
                }

                //Add as a readony text box - this will be updated when the contract is instansiated
                string txt = sel.Range.Text;

                //Delete what was there first
                sel.Delete();

                //and insert a control
                Word.ContentControl c = sel.Document.ContentControls.Add(Word.WdContentControlType.wdContentControlText);
                c.SetPlaceholderText(null, null, tbName.Text);
                c.Range.Text = "";                
                c.Title = Utility.Truncate(tbName.Text,64);
                c.Tag = "Element|" + elementid + "|" + clauseid;
                if (Utility.StyleExists(doc.Styles, "ContentControl")) c.set_DefaultTextStyle("ContentControl");
                
                c.LockContentControl = true;
                c.LockContents = false;

                //Save! important to try and keep things in sync
                try
                {
                   // Globals.ThisAddIn.ProcessingUpdate("Save Clause Template");
                   // doc.Save();
                }
                catch (Exception)
                {
                }

                //Now Reload the Tree
                Globals.ThisAddIn.ProcessingUpdate("Refresh Tree");
                Globals.ThisAddIn.GetTaskPaneControlTemplate().RefreshElements();
                
            }
            Globals.ThisAddIn.ProcessingStop("End");
            this.Hide();
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
            RefreshElementList();
            dgElements.SelectedIndex = -1;
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (btnSave.IsEnabled)
            {
                MessageBoxResult res = MessageBox.Show("Loose Changes?", "Warning", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    dgElements.SelectedIndex = -1;
                    return;
                }
            }

            //Clear the form 
            Utility.ClearForm(new Grid[] { formGrid1, formGrid2, formGrid3 });
            dgElements.SelectedIndex = -1;
            btnSave.IsEnabled = true;
            btnCancel.IsEnabled = true;
            tbName.Focus();

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

        public void Open(string elementid)
        {
            DataView dv = (DataView)dgElements.ItemsSource;
            for (int i = 0; i < dv.Table.Rows.Count; i++)
            {
                if (dv.Table.Rows[i]["Id"].ToString() == elementid)
                {
                    dgElements.SelectedIndex = i;
                }
            }
        }


    }
}
