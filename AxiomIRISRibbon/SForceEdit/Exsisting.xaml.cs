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
using AxiomIRISRibbon.Core;
using AxiomIRISRibbon.sfPartner;
using System.IO;




namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for Exsisting.xaml
    /// NEW File Added by PES
    /// </summary>
    public partial class Exsisting : RadWindow
    {

        static Microsoft.Office.Interop.Word.Application app;

        private Data _d;
        string _objname;
        string _id;
        string _name;
        private SForceEdit.SObjectDef _sDocumentObjectDef;
        public AxObject _parentObject;

        public Exsisting()
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();

            app = Globals.ThisAddIn.Application;
        }

        private void ClauseRowDoubleClick(object sender, RoutedEventArgs e)
        {
            //Open();
        }


        public void Create(string objname, string id, string name, string templatename)
        {

            _objname = objname;
            _id = id;
            _name = name;

            DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplatesFromExsisting(true));
            if (!dr.success) return;

            DataTable dt = dr.dt;
            dgTemplates.Items.Clear();

            this.dgTemplates.ItemsSource = dt.DefaultView;
            dgTemplates.Focus();
        }
        public void btnReset_Click(object sender, RoutedEventArgs e)
        {
            CNID.Text = "" ;
            AgreemntNumber.Text = "";
            DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplatesFromExsisting(true));
            if (!dr.success) return;

            DataTable dt = dr.dt;
            // dgTemplates.Items.Clear();
            this.dgTemplates.ItemsSource = dt.DefaultView;
            dgTemplates.Focus();
        }

        protected void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            string cnid;
            string agreementnumber;

            cnid = CNID.Text;
            agreementnumber = AgreemntNumber.Text;

            if (CNID.Text == "" && AgreemntNumber.Text == "")
            {
               MessageBoxResult result = MessageBox.Show("Please enter either Agreemnt Number or CNID");

            }
            else
            {

                try
                {
                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateForsearch(agreementnumber, cnid));
                    if (!dr.success) return;

                    DataTable dt = dr.dt;
                    //dgTemplates.Items.Clear();
                    this.dgTemplates.ItemsSource = dt.DefaultView;
                    dgTemplates.Focus();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, "btnSearch_Click");
                    MessageBoxResult result = MessageBox.Show("Error text here", "Caption", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
        }



        public void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        //Code PES
        protected void btnClone_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bsyIndc.IsBusy = true;
                bsyIndc.BusyContent = "Cloning ...";
              

                if ((DataRowView)dgTemplates.SelectedItem == null)
                {
                    MessageBox.Show("Select an item", "Alert");
                }
                else
                {
                        
                    double dVersionNumber = 0;
                    string strFromAgreementId, strToAgreementId, strVersionId = string.Empty, strTemplate = string.Empty;
                    strToAgreementId = _id;

                    DataRow dtr = ((DataRowView)dgTemplates.SelectedItem).Row;
                    DataRow allDr;// = new DataRow();

                    strFromAgreementId = dtr["Id"].ToString();

                    //Get version from 

                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementsForVersion(strFromAgreementId));

                    if (!dr.success) return;
                    if (dr.dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Version not avilable in source Agreement");
                    }
                    else
                    {
                        DataTable dt = dr.dt;
                         allDr = dt.NewRow();

                        foreach (DataRow r in dt.Rows)
                        {
                            strVersionId = r["Id"].ToString();
                            //   dVersionNumber = Convert.ToDouble(r["version_number__c"]);
                            strTemplate = Convert.ToString(r["Template__c"]);
                        }
                      

                        //Get version to 
                        DataReturn drTo = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementsForVersion(strToAgreementId));
                        DataTable dtrTo = drTo.dt;
                        foreach (DataRow rw in dtrTo.Rows)
                        {
                            allDr = rw;
                            //strVersionId = rw["Id"].ToString();
                            dVersionNumber = Convert.ToDouble(rw["version_number__c"]);
                            // strTemplate = Convert.ToString(r["Template__c"]);
                        }
                        double maxId;
                        if (drTo.dt.Rows.Count == 0)
                        {
                            maxId = 0;
                        }
                        else
                        {
                            maxId = Convert.ToDouble(dVersionNumber + 1);
                        }

                        string VersionName = "Version " + (maxId).ToString();
                        string VersionNumber = maxId.ToString();

                        // Create Version 0 or lower version in To
                         DataReturn drCreatev0 = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", strToAgreementId, strTemplate, VersionName, VersionNumber, allDr));
                         string newV0VersionId = drCreatev0.id;
                        // Create Version 1 or lower version +1 in To
                       maxId = Convert.ToDouble(maxId + 1);
                        VersionName = "Version " + (maxId).ToString();                        
                        VersionNumber = maxId.ToString();

                        DataReturn drCreateV1 = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", strToAgreementId, strTemplate, VersionName, VersionNumber,allDr));
                        string newV1VersionId = drCreateV1.id;

                        //Create attachments in To
                        DataReturn drVersionAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetVersionAllAttachments(strVersionId));
                        if (!drVersionAttachemnts.success) return;
                        DataTable dtAttachments = drVersionAttachemnts.dt;

                        if (dtAttachments.Rows.Count == 0)
                        {
                            MessageBox.Show("Attachments not avilable in source Version");
                        }
                        else
                        {
                            string filename = "";
                            foreach (DataRow rw in dtAttachments.Rows)
                            {
                                filename = rw["Name"].ToString();
                                string body = rw["body"].ToString();
                                _d.saveAttachmentstoSF(newV0VersionId, filename, body);
                                _d.saveAttachmentstoSF(newV1VersionId, filename, body);
                            }

                            //Get Attachments
                            DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetAllAttachments(newV1VersionId));
                            if (!drAttachemnts.success) return;
                            DataTable dtAllAttachments = drAttachemnts.dt;

                            //Open attachment with compare screeen
                            OpenAttachment(dtAllAttachments, newV1VersionId, strToAgreementId, strTemplate, VersionName, VersionNumber);
                       //     _sDocumentObjectDef = new SForceEdit.SObjectDef("Version__c");
                            Globals.Ribbons.Ribbon1.CloseWindows();
                        }
                    }
                }
                bsyIndc.IsBusy = false;
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }
       
        private  void OpenAttachment(DataTable dt,string versionid, string matterid, string templateid, string versionName, string versionNumber)
        {
            try
            {
                    var res = from row in dt.AsEnumerable()
                              where 
                              (row.Field<string>("Name").Contains(".doc") ||
                              row.Field<string>("ContentType").Contains("msword"))
                              select row;
                    if (res.Count() > 1)
                    {

                        AttachmentsView attTemp = new AttachmentsView();
                        attTemp.Create(dt,versionid, matterid, templateid, versionName, versionNumber);
                        attTemp.Show();
                    }
                    else
                    {

                        string attachmentid;
                        foreach (DataRow rw in dt.Rows)
                        {
                            if (rw["Name"].ToString().Contains(".doc"))
                            {
                                byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());

                                File.WriteAllBytes(filename, toBytes);
                                // _source = app.Documents.Open(filename);


                                Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(filename);
                                //     Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                                //     doc.Activate();

                                attachmentid = rw["Id"].ToString();

                                //Right Panel
                                System.Windows.Forms.Integration.ElementHost elHost = new System.Windows.Forms.Integration.ElementHost();
                                SForceEdit.CompareSideBar csb = new SForceEdit.CompareSideBar();
                                csb.Create(filename, versionid, matterid, templateid, versionName, versionNumber, attachmentid);

                                elHost.Child = csb;
                                elHost.Dock = System.Windows.Forms.DockStyle.Fill;
                                System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
                                u.Controls.Add(elHost);
                                Microsoft.Office.Tools.CustomTaskPane taskPaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(u, "Axiom IRIS Compare", doc.ActiveWindow);
                                taskPaneValue.Visible = true;
                                taskPaneValue.Width = 400;
                            }
                        }
                    }
            }
            catch (Exception ex) { Logger.Log(ex, "OpenAttachment"); }
           
        }

        

        //End Code PES
    }
}

