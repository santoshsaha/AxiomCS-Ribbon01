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

using AxiomIRISRibbon.Core;
using System.IO;


namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for SForceEditSideBar.xaml
    /// New File added by PES
    /// </summary>
    public partial class CompareSideBar : UserControl
    {

        private Data _d;
        static Microsoft.Office.Interop.Word.Application app;
        static Word.Document _source;
        private string _fileName;

        private string _matterid;
        private string _versionid;
        private string _templateid;
        private string _versionName;
        private string _versionNumber;
        private string _attachmentid;
        private Word.Document _doc;

        public CompareSideBar()
        {
           InitializeComponent();
           AxiomIRISRibbon.Utility.setTheme(this);

           _d = Globals.ThisAddIn.getData();

           LoadTemplatesDLL();
        
        }
        public void Create(string filename, string versionid, string matterid, string templateid, string versionName, string versionNumber, string attachmentid)
        {
            _fileName = filename;
            _matterid = matterid;
               _versionid = versionid;
               _templateid = templateid;
               _versionName = versionName;
               _versionNumber = versionNumber;
               _attachmentid = attachmentid;
        
        }
        private void LoadTemplatesDLL()
        {
            try
            {

                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplates(true));
                if (!dr.success) return;

                DataTable dt = dr.dt;
                cbTemplates.Items.Clear();

                RadComboBoxItem i;

                // RadComboBoxItem selected = null;
                foreach (DataRow r in dt.Rows)
                {
                    i = new RadComboBoxItem();
                    i.Tag = r["Id"].ToString() ;
                    i.Content = r["Name"].ToString();
                    this.cbTemplates.Items.Add(i);

                }

            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }


        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;

                if (this.cbTemplates.SelectedItem != null)
                {

                    string TemplateId = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Tag.ToString();
                    string TemplateName = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Content.ToString();
                   // Microsoft.Office.Interop.Word.Document tempDoc;

                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(TemplateId));
                    if (!drAttachemnts.success) return;

                     
                    DataTable dtAttachments = drAttachemnts.dt;
                    string file2name = "";
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                        file2name = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                        File.WriteAllBytes(file2name, toBytes);
                     //   _source = app.Documents.Open(filename);
                        
                        
                    }      

                    Microsoft.Office.Interop.Word.Document tempDoc1;
                    Microsoft.Office.Interop.Word.Document tempDoc2;
                    Microsoft.Office.Interop.Word.Application app = Globals.ThisAddIn.Application;

                 

                    object newFilenameObject2 = file2name;
                    tempDoc2 = app.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                    object newFilenameObject1 = _fileName;
                    tempDoc1 = app.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing);
                    //Compare
                    Globals.ThisAddIn.AddDocId(tempDoc1, "Compare", "");

                    object o = tempDoc2;
                    tempDoc1.Windows.CompareSideBySideWith(ref o);

                    Globals.Ribbons.Ribbon1.CloseWindows();

                }
                else
                {
                    MessageBox.Show("Select a template");

                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }
        public bool SaveContract(bool ForceSave, bool SaveDoc)
        {
            string strFileAttached = _fileName;
            //Save the Contract    
            Globals.ThisAddIn.RemoveSaveHandler(); // remove the save handler to stop the save calling the save etc.
     
            Globals.ThisAddIn.ProcessingStart("Save Contract");
            DataReturn dr;
          _doc = Globals.ThisAddIn.Application.ActiveDocument;

            dr = AxiomIRISRibbon.Utility.HandleData(_d.SaveVersion(_versionid, _matterid, _templateid, _versionName, _versionNumber));
            if (!dr.success) return false;
            _versionid = dr.id;

            if (SaveDoc)
            {        

                //Save the file as an attachment
                //save this to a scratch file
                Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
             //   string filename = AxiomIRISRibbon.Utility.SaveTempFile(_versionid);
                _doc.SaveAs2(FileName: strFileAttached, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                //Save a copy!
                Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                string filenamecopy =  AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "X");
                Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(strFileAttached, Visible: false);
                dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                docclose.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                //Now save the file - change this to always save as the version name

                Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                string vfilename = _versionName.Replace(" ", "_") + ".docx";         
                dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(_attachmentid, vfilename, filenamecopy));
           
            }
            Globals.ThisAddIn.AddSaveHandler(); // add it back in
            Globals.ThisAddIn.ProcessingStop("End");
            return true;
        }


    }
}
