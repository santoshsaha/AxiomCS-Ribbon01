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
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Data;

namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for NewFromTemplate.xaml
    /// </summary>
    public partial class NewFromTemplate : RadWindow
    {

        private Data _d;
        string _objname;
        string _id;
        string _name;

        public NewFromTemplate()
        {
            InitializeComponent();            
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


        public void Create(string objname, string id,string name, string templatename)
        {

            _objname = objname;
            _id = id;
            _name = name;

            DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplates(true));
            if (!dr.success) return;

            DataTable dt = dr.dt;
            cbTemplates.Items.Clear();

            RadComboBoxItem i;

            RadComboBoxItem selected = null;
            foreach (DataRow r in dt.Rows)
            {
                i = new RadComboBoxItem();

                i.Tag = r["Id"].ToString() + "|" + r["PlaybookLink__c"].ToString();
                i.Content = r["Name"].ToString();
                this.cbTemplates.Items.Add(i);

                if (r["Name"].ToString().ToLower() == templatename.Trim().ToLower()) selected = i;
            }

            // if we have a match then select it
            if (selected != null)
            {
                this.cbTemplates.SelectedItem = selected;
            }

            


        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            // create a new Template
            Globals.Ribbons.Ribbon1.CloseWindows();



            if(this.cbTemplates.SelectedItem!=null)
            {

                string tag = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Tag.ToString();
                string[] atag = tag.Split('|');

                string TemplateId = atag[0];
                string TemplateName = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Content.ToString();
                string TemplatePlaybookLink = "";
                if (atag.Length > 1)
                {
                    TemplatePlaybookLink = atag[1];
                }

                Globals.ThisAddIn.ProcessingStart("Create New Contract");
                this.Close(); //close first or it will switch back to this doc
                //Get the template details 
                Globals.ThisAddIn.ProcessingUpdate("Get the Template File");
            

                string filename;
                filename = AxiomIRISRibbon.Utility.SaveTempFile(TemplateId);
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateFile(TemplateId, filename));
                if (!dr.success) return;

                //Open the doc
                Word.Document doc = null;
                try
                {
                    doc = Globals.ThisAddIn.Application.Documents.Open(filename);
                    Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                }
                catch (Exception)
                {
                    Globals.ThisAddIn.ProcessingStop("Stop");
                }

                //If this is the template then change
                Globals.ThisAddIn.AddDocId(doc, "Contract", "");

                Globals.ThisAddIn.AddContractContentControlHandler(doc);

                //if the action bar isn't open then open it
                Globals.ThisAddIn.ShowTaskPane(true);

                ContractEdit.SForceEditSideBar2 u = Globals.ThisAddIn.GetTaskPaneControlContract();
                if (u != null)
                {
                    //don't get it to set the defaults if there are going to be values to load
                    u.BuildSideBarNewVersion(TemplateId, TemplateName, TemplatePlaybookLink, _id, _name);
                }

                //Scroll to the top
                Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = true;
                Globals.ThisAddIn.Application.ActiveWindow.VerticalPercentScrolled = 0;
                Globals.ThisAddIn.ProcessingStop("Stop");

                if (doc != null) doc.Activate();

            }
        }


    }
}
