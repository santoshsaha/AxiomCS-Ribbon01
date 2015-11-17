using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Data;
using Telerik.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AxiomIRISRibbon.Core;


namespace AxiomIRISRibbon
{
    public partial class ThisAddIn
    {

        //Globals
        private Data _d;
        bool _showTaskPane;
        private Processing _p;

        private Login _ucLogin;
        private LoginSSO _ucLoginSSO;
        private About _ucAbout;
        private AboutReleaseNotes _ucAboutReleaseNotes;
        private Settings _ucLocalSettings;
        private Template _ucTemplate;
        private Clause _ucClause;
        private Element _ucElement;
        private Contract _ucContract;
        private Concept _ucConcept;

        private Dictionary<string,Edit> _editWindows;
        private Dictionary<string, Edit> _editZoomWindows;
        

        private SForceEdit.Settings _settings;
        private LocalSettings _localSettings;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
            // get the local settings and set the theme
            Logger.Log("get the local settings and set the theme");
            _localSettings = new LocalSettings();
            this.SetTheme();


            //Create the Salesforce connection
            _d = new Data();
            _p = new Processing();
            _p.Hide();


            if (_localSettings.SSOLogin)
            {
                // switch off normal login
                Globals.Ribbons.Ribbon1.btnLogin.Visible = false;

                if (_localSettings.ShowAllLogins)
                {
                    Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = false;
                    Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = true;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = true;
                    Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = false;
                }
            }
            else
            {
                Globals.Ribbons.Ribbon1.btnLogin.Visible = true;
                Globals.Ribbons.Ribbon1.btnLoginSSO.Visible = false;
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Visible = false;
            }

            // set the name of the Ribbon button depending on the selected instances
            if (_localSettings.Inst != null && _localSettings.Inst != LocalSettings.Instances.Prod)
            {
                Globals.Ribbons.Ribbon1.btnLoginSSO.Label = _localSettings.Inst.ToString();
                Globals.Ribbons.Ribbon1.sbtnLoginSSO.Label = _localSettings.Inst.ToString();
            }

            

            if (_localSettings.Debug)
            {
                Globals.Ribbons.Ribbon1.gpDebug.Visible = true;

                // switch on normal login
                Globals.Ribbons.Ribbon1.btnLogin.Visible = true;
                Globals.Ribbons.Ribbon1.btnLoginSSO.Label = "SSO";

            }
            else
            {
                Globals.Ribbons.Ribbon1.gpDebug.Visible = false;
            }




            // ----------------------------------------------------------------------------------------------------------------------------------------------------
            // Change Nov : Auto Login
            bool autologin = false;

            if (autologin && _localSettings.Debug)
            {

                this.SetRibbon();                   
                //Globals.ThisAddIn.ProcessingStart("AutoLogin - remember to switch off!");

                // This can be used for testing so you don't have to login every time
                // set the autologin about to true
                // then add the details to the login call - should be username, password, sforce token, sforce url, login description to show in the about 
                string rtn = _d.Login("santosh.saha@cs.com.rksb1", "pass@word1", "LGZ0rTkNnuksEetJr1vrG0YS", "https://test.salesforce.com", "AutoLogin - Sales");

                if (rtn == "")
                {
                    
                    // get the settings
                    bool gotSettings = true;

                    DataReturn settings = _d.GetStaticResource("RibbonSettings");
                    if (!settings.success || settings.strRtn == "")
                    {
                        // get the default settings                      
                        var uri = new Uri("pack://application:,,,/AxiomIRISRibbon;component/Resources/Settings.json");
                        byte[] bjson = AxiomIRISRibbon.Properties.Resources.Settings;
                        string sjson = Encoding.Default.GetString(bjson, 0, bjson.Length - 1);
                        try
                        {
                            _settings = new SForceEdit.Settings(sjson);
                            gotSettings = true;
                        }
                        catch (Exception eSet)
                        {
                            System.Windows.MessageBoxResult rslt = System.Windows.MessageBox.Show("Settings did not load : " + eSet.Message, "Axiom IRIS", System.Windows.MessageBoxButton.OK);
                        }
                    }
                    else
                    {
                        try
                        {
                            _settings = new SForceEdit.Settings(settings.strRtn);
                            gotSettings = true;
                        }
                        catch (Exception eSet)
                        {
                            System.Windows.MessageBoxResult rslt = System.Windows.MessageBox.Show("Settings did not load : " + eSet.Message, "Axiom IRIS", System.Windows.MessageBoxButton.OK);
                        }
                        
                    }

                    if (!gotSettings) return;
                    
                    DataReturn dr = Utility.HandleData(_d.LoadDefinitions());
                    Globals.Ribbons.Ribbon1.LoginOK();
                    Globals.Ribbons.Ribbon1.gpDebug.Visible = true;
                    
                 
                }
                //----------------------------------
                                
            }
            else
            {
               
            }
           
             Globals.ThisAddIn.ProcessingStop("LoggedIN");
           

            // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/


             _editWindows = new Dictionary<string, Edit>();
             _editZoomWindows = new Dictionary<string, Edit>();

            //Add in the Save handler
            this.Application.DocumentBeforeSave +=
             new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);

            //Add a loader
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);

            //add a handler to tidy up sidebars
            this.Application.DocumentChange += Application_DocumentChange;

        }

        void Application_DocumentChange()
        {
            ClearUpTaskPanes();
        }


        public SForceEdit.Settings GetSettings(){
            return _settings;
        }

        public void setSettings(SForceEdit.Settings s)
        {
            _settings = s;
        }

        public string GetSettings(string sObject, string key){
            return _settings.GetSetting(sObject,key);
        }

        public string GetSettings(string sObject, string key,string subkey)
        {
            return _settings.GetSetting(sObject, key, subkey);
        }

        public LocalSettings GetLocalSettings()
        {
            return _localSettings;
        }

        public  void SaveLocalSettings(LocalSettings s)
        {
            _localSettings = s;
        }

        public void AddSaveHandler()
        {
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);

        }

        public void RemoveSaveHandler(){
            this.Application.DocumentBeforeSave -= new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        void Application_DocumentOpen(Word.Document Doc)
        {

            // Close out any toolbars
            ClearUpTaskPanes();

            string tag = GetCurrentAxiomDocProp();
            if (tag != "" && tag.Contains("|"))
            {
                if (tag.Split('|')[0] == "ExportContract")
                {
                    System.Windows.MessageBoxResult rslt = System.Windows.MessageBox.Show("Would you like to import this version of the doc?","Axiom IRIS",System.Windows.MessageBoxButton.OKCancel);

                    if (rslt == MessageBoxResult.OK)
                    {
                        // Modified Contract Open procedure!
                        Globals.ThisAddIn.OpenContract().OpenClauseFromNegotiatedDoc(tag.Split('|')[1]);
                    }
                }

                if (tag.Split('|')[0] == "ExportTemplate")
                {

                    string currentinstance = _d.GetInstanceInfo();
                    string prompt = "Would you like to import this template?\n\nYour are currently logged into: " + currentinstance;

                    System.Windows.MessageBoxResult rslt = System.Windows.MessageBox.Show(prompt, "Axiom IRIS", System.Windows.MessageBoxButton.OKCancel);

                    if (rslt == MessageBoxResult.OK)
                    {
                        // Load the Template Import
                        Globals.ThisAddIn.OpenTemplate().OpenImportTemplate();
                    }
                }

            }
            
        }

        public bool getDebug()
        {
            return _localSettings.Debug;
        }

        public void SetRibbon(){
            Globals.Ribbons.Ribbon1.Activate();
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        public Data getData()
        {
            return _d;
        }

        public SForceEdit.Settings getSettings()
        {
            return _settings;
        }

        public void ProcessingStart(string t)
        {
            _p.Start(t);
            return;
        }

        public void ProcessingUpdate(string t)
        {
            _p.Update(t);
            return;
        }

        public void ProcessingStop(string t)
        {
            _p.Stop(t);
            return;
        }

        public void HideWindows(){
            if (_ucLogin != null) _ucLogin.Close();
            if (_ucLoginSSO != null) _ucLoginSSO.Close();
            if (_ucAbout != null) _ucAbout.Close();
            if (_ucClause != null) _ucClause.Hide();
            if (_ucTemplate != null) _ucTemplate.Hide();
            if (_ucElement != null) _ucElement.Hide();
            if (_ucContract != null) _ucContract.Hide();
            if (_ucConcept != null) _ucConcept.Hide();
        }

        public Login OpenLogin()
        {
            HideWindows();
            if (_ucLogin == null)
            {
                _ucLogin = new Login();
            }
            _ucLogin.Show();
            _ucLogin.BringToFront();
            return _ucLogin;
        }


        public LoginSSO OpenLoginSSO()
        {
            HideWindows();
            if (_ucLoginSSO == null)
            {
                _ucLoginSSO = new LoginSSO();
            }
            _ucLoginSSO.Show();
            _ucLoginSSO.BringToFront();
            return _ucLoginSSO;
        }

        public About OpenAbout()
        {
            HideWindows();
            if (_ucAbout == null)
            {
                _ucAbout = new About();
            }
            _ucAbout.Show();
            _ucAbout.BringToFront();
            return _ucAbout;
        }

        public AboutReleaseNotes OpenAboutReleaseNotes()
        {
            if (_ucAboutReleaseNotes == null)
            {
                _ucAboutReleaseNotes = new AboutReleaseNotes();
            }
            _ucAboutReleaseNotes.Show();
            _ucAboutReleaseNotes.BringToFront();
            return _ucAboutReleaseNotes;
        }

        public Settings OpenLocalSettings()
        {
            HideWindows();
            if (_ucLocalSettings == null)
            {
                _ucLocalSettings = new Settings();
            }
            _ucLocalSettings.Show();
            _ucLocalSettings.BringToFront();
            return _ucLocalSettings;
        }

        public Template OpenTemplate()
        {
            HideWindows();
            if (_ucTemplate == null)
            {
                _ucTemplate = new Template();
            }
            else
            {
                _ucTemplate.Refresh();
            }

            _ucTemplate.Show();
            _ucTemplate.Activate();
            return _ucTemplate;
        }
        public Clause OpenClause(bool show,bool refresh)
        {
            HideWindows();
            if (_ucClause == null)
            {
                _ucClause = new Clause(refresh);
            }
            else
            {
                if(refresh) _ucClause.RefreshIfNotLoaded();
            }
            if(show) _ucClause.Show();
            return _ucClause;
        }
        public Element OpenElement()
        {
            HideWindows();
            if (_ucElement == null)
            {
                _ucElement = new Element();
            }
            else
            {
                _ucElement.Refresh();
            }
            _ucElement.Show();
            _ucElement.Activate();
            return _ucElement;
        }
        public Contract OpenContract()
        {
            HideWindows();
            if (_ucContract == null)
            {
                _ucContract = new Contract();
            }
            else
            {
                _ucContract.Refresh();
            }
            _ucContract.Show();
            _ucContract.Activate();
            return _ucContract;
        }
        public Concept OpenConcept()
        {
            HideWindows();
            if (_ucConcept == null)
            {
                _ucConcept = new Concept();
            }
            else
            {
                _ucConcept.Refresh();
            }
            _ucConcept.Show();
            _ucConcept.Activate();
            _ucConcept._editmode = "";
            return _ucConcept;
        }

        //------------------ Task Pane Stuff

        //Manage multiple TaskPanes - each doc has its own have to manage them -----
        public void ShowTaskPane(bool show)
        {
            ClearUpTaskPanes();

            _showTaskPane = show;
            //If we've to show then check we haven't got one already, if not create one
            bool found = false;
            if (_showTaskPane)
            {
                try
                {
                    foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                    {
                        try
                        { // Nov 5
                            if ((ctp.Title == "Axiom IRIS Template" || ctp.Title == "Axiom IRIS Contract" || ctp.Title == "Axiom IRIS Compare") && ctp.Window == this.Application.ActiveWindow)
                            {
                                ctp.Visible = true;
                                found = true;

                                //Check we have the right one
                                if ((ctp.Title == "Axiom IRIS Template" && !isTemplate()) || (ctp.Title == "Axiom IRIS Contract" && !isContract() && !isUnAttachedContract()))
                                {
                                    this.CustomTaskPanes.Remove(ctp);
                                    AddAwesomeAxiomTaskPane(this.Application.ActiveDocument);
                                }


                            }
                        }
                        catch
                        {
                        }

                    }
                }
                catch
                {
                }

                if (!found && this.Application.Documents.Count > 0) AddAwesomeAxiomTaskPane(this.Application.ActiveDocument);
            }
            else
            {

                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                {
                    try
                    {
                        if ((ctp.Title == "Axiom IRIS Template" || ctp.Title == "Axiom IRIS Contract") && ctp.Window == this.Application.ActiveWindow)
                        {
                            ctp.Visible = false;
                        }
                    }
                    catch
                    {
                    }
                }
            }

        }

        private void ClearUpTaskPanes()
        {
            for (int i = this.CustomTaskPanes.Count; i > 0; i--)
            {
                Microsoft.Office.Tools.CustomTaskPane ctp = this.CustomTaskPanes[i - 1];
                try
                {
                    if (ctp != null && ctp.Window == null)
                    {
                        this.CustomTaskPanes.Remove(ctp);
                    }
                }
                catch
                {
                }
            }

        }
       
        private void AddAwesomeAxiomTaskPane(Word.Document doc)
        {
            this.SetRibbon();

            // WPF Form
            if (isTemplate() || isExportTemplate() || isClause())
            {

                System.Windows.Forms.Integration.ElementHost elHost = new System.Windows.Forms.Integration.ElementHost();
                TemplateEdit.TEditSidebar tsb = new TemplateEdit.TEditSidebar(doc);
                elHost.Child = tsb;
                elHost.Dock = System.Windows.Forms.DockStyle.Fill;
                System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
                u.Controls.Add(elHost);
                Microsoft.Office.Tools.CustomTaskPane taskPaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(u, "Axiom IRIS Template", doc.ActiveWindow);
                taskPaneValue.Visible = true;
                taskPaneValue.Width = 300;
                taskPaneValue.VisibleChanged += new EventHandler(taskPaneValue_VisibleChanged);
                
            }
            else if (isContract() || isUnAttachedContract())
            {
                
                ElementHost elHost = new ElementHost();
                ContractEdit.SForceEditSideBar2 csb = new ContractEdit.SForceEditSideBar2();
                elHost.Child = csb;
                elHost.Dock = DockStyle.Fill;
                System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
                u.Controls.Add(elHost);
                Microsoft.Office.Tools.CustomTaskPane taskPaneValue = this.CustomTaskPanes.Add(u, "Axiom IRIS Contract", doc.ActiveWindow);
                taskPaneValue.Visible = true;
                taskPaneValue.Width = 400;
                taskPaneValue.VisibleChanged += new EventHandler(taskPaneValue_VisibleChanged);

            } 

        }


        public void ShowTaskPaneSFEdit(Word.Document doc, bool show,string Id,string FileName,string ParentType,string ParentId)
        {
            ClearUpTaskPanes();

            _showTaskPane = show;
            //If we've to show then check we haven't got one already, if not create one
            bool found = false;
            if (_showTaskPane)
            {
                try
                {
                    foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                    {
                        try
                        {
                            if ((ctp.Title == "Axiom IRIS Edit") && ctp.Window == this.Application.ActiveWindow)
                            {
                                ctp.Visible = true;
                                found = true;

                                    this.CustomTaskPanes.Remove(ctp);
                                    AddAwesomeAxiomTaskPaneSFEdit(doc, Id, FileName, ParentType, ParentId);
  
                            }
                        }
                        catch
                        {
                        }

                    }
                }
                catch
                {
                }

                if (!found && this.Application.Documents.Count > 0) AddAwesomeAxiomTaskPaneSFEdit(doc, Id, FileName, ParentType, ParentId);
            }
            else
            {

                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                {
                    try
                    {
                        if ((ctp.Title == "Axiom IRIS Edit") && ctp.Window == this.Application.ActiveWindow)
                        {
                            ctp.Visible = false;
                        }
                    }
                    catch
                    {
                    }
                }
            }

        }

        private void AddAwesomeAxiomTaskPaneSFEdit(Word.Document doc, string Id, string FileName, string ParentType, string ParentId)
        {
            this.SetRibbon();
            System.Windows.Forms.Integration.ElementHost elHost = new System.Windows.Forms.Integration.ElementHost();
            ContractEdit.SForceEditSideBar2 ssb = new ContractEdit.SForceEditSideBar2(Id, FileName, ParentType, ParentId);
            elHost.Child = ssb;
            elHost.Dock = System.Windows.Forms.DockStyle.Fill;
            System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
            u.Controls.Add(elHost);
            Microsoft.Office.Tools.CustomTaskPane taskPaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(u, "Axiom IRIS Edit", doc.ActiveWindow);
            taskPaneValue.Visible = true;
            taskPaneValue.Width = 300;
        }


        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            Microsoft.Office.Tools.CustomTaskPane ctp = (Microsoft.Office.Tools.CustomTaskPane)sender;
            if (ctp.Visible != this._showTaskPane)
            {
                this._showTaskPane = ctp.Visible;
            }
        }


        public Microsoft.Office.Tools.CustomTaskPane GetTaskPane()
        {
            return GetTaskPane(this.Application.ActiveDocument);
        }

        public Microsoft.Office.Tools.CustomTaskPane GetTaskPane(Word.Document doc)
        {
            try
            {
                foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
                {
                    try
                    {
                        if ((ctp.Title == "Axiom IRIS Template" || ctp.Title == "Axiom IRIS Contract" ) && ctp.Window == doc.ActiveWindow)
                        {
                            return ctp;
                        }
                        else if (ctp.Title == "Axiom IRIS Compare")
                        {
                            return ctp;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex, "GetTaskPane-CustomTaskPanes");
                    }

                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "GetTaskPane");
            }
            return null;
        }




        public TemplateEdit.TEditSidebar GetTaskPaneControlTemplate()
        {
            return GetTaskPaneControlTemplate(this.Application.ActiveDocument);
        }

        public ContractEdit.SForceEditSideBar2 GetTaskPaneControlContract()
       {
           return GetTaskPaneControlContract(this.Application.ActiveDocument);
        }

       public TemplateEdit.TEditSidebar GetTaskPaneControlTemplate(Word.Document doc)
        {
            Microsoft.Office.Tools.CustomTaskPane ctp = GetTaskPane(doc);
            if (ctp != null)
            {
                System.Windows.Forms.UserControl u = ctp.Control;
                ElementHost elHost = (ElementHost)u.Controls[0];
                if (elHost.Child.GetType().ToString() == "AxiomIRISRibbon.TemplateEdit.TEditSidebar") return ((TemplateEdit.TEditSidebar)elHost.Child);
          
            }
            return null;
        }


       public ContractEdit.SForceEditSideBar2 GetTaskPaneControlContract(Word.Document doc)
        {
            Microsoft.Office.Tools.CustomTaskPane ctp = GetTaskPane(doc);
            if (ctp != null)
            {
                System.Windows.Forms.UserControl u = ctp.Control;
                ElementHost elHost = (ElementHost)u.Controls[0];
                if (elHost.Child.GetType().ToString() == "AxiomIRISRibbon.ContractEdit.SForceEditSideBar2") return ((ContractEdit.SForceEditSideBar2)elHost.Child);
              
            }
            return null;
        }
        
        //NEW PES
       public SForceEdit.CompareSideBar GetTaskPaneControlCompare()
       {
           Word.Document doc = this.Application.ActiveDocument;
           Microsoft.Office.Tools.CustomTaskPane ctp = GetTaskPane(doc);
           if (ctp != null)
           {
               System.Windows.Forms.UserControl u = ctp.Control;
               ElementHost elHost = (ElementHost)u.Controls[0];
              if (elHost.Child.GetType().ToString() == "AxiomIRISRibbon.SForceEdit.CompareSideBar") return ((SForceEdit.CompareSideBar)elHost.Child);
           }
           return null;
       }
        //END PES

        // Update all the data controls on all the task panes - could get round doing this if we tied to a datatable
        // but this will do for now!
        // Ok this takes a while when using salesforce so try and avoid!
        public void RefreshAllTaskPanes()
        {
            foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
            {
                try
                {
                    if (ctp.Title == "Axiom IRIS Template")
                    {
                        System.Windows.Forms.UserControl u = ctp.Control;
                        ElementHost elHost = (ElementHost)u.Controls[0];
                        ((TemplateEdit.TEditSidebar)elHost.Child).Refresh();
                    }
                    else
                    {

                        System.Windows.Forms.UserControl u = ctp.Control;
                        ElementHost elHost = (ElementHost)u.Controls[0];
                        ((ContractEdit.SForceEditSideBar2)elHost.Child).Refresh();

                    }
                }
                catch (Exception e)
                {
                }
            }
        }

        // OK - hitting salesforce is expensive - so search the tree and update the xml
        public void RefreshAllTaskPanesWithClause(string Id,string Xml)
        {
            foreach (Microsoft.Office.Tools.CustomTaskPane ctp in this.CustomTaskPanes)
            {
                try
                {
                    if (ctp.Title == "Axiom IRIS Template")
                    {
                        System.Windows.Forms.UserControl u = ctp.Control;
                        ElementHost elHost = (ElementHost)u.Controls[0];
                        ((TemplateEdit.TEditSidebar)elHost.Child).RefreshMatchClause(Id, Xml);
                    }
                    else
                    {
                        System.Windows.Forms.UserControl u = ctp.Control;
                        ElementHost elHost = (ElementHost)u.Controls[0];
                        ((ContractEdit.SForceEditSideBar2)elHost.Child).Refresh(); ;
                    }
                }
                catch (Exception e)
                {
                }
            }
        }

        public void RefreshTaskPane()
        {

            Microsoft.Office.Tools.CustomTaskPane ctp = GetTaskPane(this.Application.ActiveDocument);
            if (ctp != null)
            {
                if (ctp.Title == "Axiom IRIS Template")
                {
                    System.Windows.Forms.UserControl u = ctp.Control;
                    ElementHost elHost = (ElementHost)u.Controls[0];
                    ((TemplateEdit.TEditSidebar)elHost.Child).Refresh();

                }
                else
                {
                    System.Windows.Forms.UserControl u = ctp.Control;
                    ElementHost elHost = (ElementHost)u.Controls[0];
                    ((ContractEdit.SForceEditSideBar2)elHost.Child).Refresh(); ;
                }
            }
        }
        
        // -------------- End of Task Pane Stuff



        //--------------- Word Doc stuff
        
        //Wire up the doc to handle entering Content Controls 
        public void AddContentControlHandler(Word.Document doc)
        {
            Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
            vstoDoc.ContentControlOnEnter += new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter);
        }

        void doc_ContentControlOnEnter(Word.ContentControl cc)
        {
            if (!isContract())
            {
                string tag = cc.Tag;
                if (tag != null && tag != "" && cc.Tag.Contains('|'))
                {
                    //Get the toolbar
                    TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate();

                    if (tsb != null)
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept")
                        {
                            tsb.SelectConcept(taga[1]);
                        }
                        else if (taga[0] == "Element")
                        {

                            // small fix - make sure content control is editable
                            // but not deletable - this will let the user modify the format
                            cc.LockContentControl = true;
                            cc.LockContents = false;

                            tsb.SelectElement(taga[1]);
                        }
                    }
                }
            }
            
        }

        //Wire up the Contract doc to handle exiting Content Controls 
        public void AddContractContentControlHandler(Word.Document doc)
        {
            if (doc != null)
            {
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                vstoDoc.ContentControlOnEnter += new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter);
                vstoDoc.ContentControlOnExit += new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);
            }
        }

        void vstoDoc_ContentControlOnExit(Word.ContentControl cc, ref bool Cancel)
        {
            if (isContract())
            {
                string tag = cc.Tag;
                if (tag != null && tag != "" && cc.Tag.Contains('|'))
                {
                    //Get the toolbar
                    ContractEdit.SForceEditSideBar2 csb = Globals.ThisAddIn.GetTaskPaneControlContract();

                    string[] taga = cc.Tag.Split('|');
                    if (taga[0] == "Element")
                    {
                        //Update the value in the forms
                        csb.UpdateElement(Convert.ToString(taga[1]), cc.Range.Text,"");
                    }
                }
            }
        }


        public bool isTemplate()
        {
            return isTemplate(Application.ActiveDocument);
        }

        public bool isTemplate(Word.Document doc)
        {
            bool out1 = false;
            string prop = ReadDocumentProperty(doc, "Axiom");
            if (prop != null && prop.Contains('|'))
            {
                string[] propa = prop.Split('|');
                if (propa[0] == "ContractTemplate") out1 = true;
            }
            return out1;
        }

        public bool isExportTemplate()
        {
            return isExportTemplate(Application.ActiveDocument);
        }

        public bool isExportTemplate(Word.Document doc)
        {
            bool out1 = false;
            string prop = ReadDocumentProperty(doc, "Axiom");
            if (prop != null && prop.Contains('|'))
            {
                string[] propa = prop.Split('|');
                if (propa[0] == "ExportTemplate") out1 = true;
            }
            return out1;
        }

        public bool isClause()
        {
            return isClause(Application.ActiveDocument);
        }

        public bool isClause(Word.Document doc)
        {
            bool out1 = false;
            string prop = ReadDocumentProperty(doc, "Axiom");
            if (prop != null && prop.Contains('|'))
            {
                string[] propa = prop.Split('|');
                if (propa[0] == "ClauseTemplate") out1 = true;
            }
            return out1;
        }

        public bool isContract()
        {
            return isContract(Application.ActiveDocument);
        }

        public bool isContract(Word.Document doc)
        {
            bool out1 = false;
            try
            {
                string prop = ReadDocumentProperty(doc, "Axiom");
                if (prop != null && prop.Contains('|'))
                {
                    string[] propa = prop.Split('|');
                    if (propa[0] == "Contract") out1 = true;
                }
            }
            catch (Exception)
            {
            }
            return out1;
        }

        public bool isUnAttachedContract()
        {
            return isUnAttachedContract(Application.ActiveDocument);
        }

        public bool isUnAttachedContract(Word.Document doc)
        {
            bool out1 = false;
            try
            {
                string prop = ReadDocumentProperty(doc, "Axiom");
                if (prop != null && prop.Contains('|'))
                {
                    string[] propa = prop.Split('|');
                    if (propa[0] == "UAContract") out1 = true;
                }
            }
            catch (Exception)
            {
            }
            return out1;
        }


        public string GetDocId(Word.Document doc)
        {
            string out1 = "";
            string prop = ReadDocumentProperty(doc, "Axiom");
            if (prop != null && prop.Contains('|'))
            {
                string[] propa = prop.Split('|');
                out1 = propa[1];
            }
            return out1;
        }

        public void AddDocId(Word.Document doc, string type, string id)
        {
            string prop = ReadDocumentProperty(doc, "Axiom");
            if (prop != null && prop != "")
            {
                //Delete and add
                Office.DocumentProperties properties = (Office.DocumentProperties)doc.CustomDocumentProperties;

                properties["Axiom"].Delete();
                properties.Add("Axiom", false, Office.MsoDocProperties.msoPropertyTypeString, type + "|" + id);

            }
            else
            {
                Office.DocumentProperties properties = (Office.DocumentProperties)doc.CustomDocumentProperties;
                properties.Add("Axiom", false, Office.MsoDocProperties.msoPropertyTypeString, type + "|" + id);
            }
            return;
        }

        public void DeleteDocId(Word.Document doc)
        {
            string prop = ReadDocumentProperty(doc, "Axiom");
            if (prop != null && prop != "")
            {
                Office.DocumentProperties properties = (Office.DocumentProperties)doc.CustomDocumentProperties;
                //Delete
                properties["Axiom"].Delete();                
            }

            return;
        }

        public string GetCurrentDocId()
        {
            string out1 = "";
            if (this.Application.Documents.Count > -1)
            {
                return GetDocId(this.Application.ActiveDocument);

            }
            return out1;
        }

        public string GetCurrentAxiomDocProp()
        {
            string out1 = "";
            try
            {
                if (this.Application.Documents.Count > 0)
                {
                    string prop = ReadDocumentProperty(this.Application.ActiveDocument, "Axiom");
                    if (prop != null) out1 = prop;
                }
            }
            catch (Exception)
            {

            }
            return out1;
        }

        private string ReadDocumentProperty(Word.Document doc, string propertyName)
        {
            Office.DocumentProperties properties;
            if (doc.CustomDocumentProperties != null)
            {
                properties = (Office.DocumentProperties)doc.CustomDocumentProperties;

                if (properties != null)
                {
                    foreach (Office.DocumentProperty prop in properties)
                    {
                        if (prop.Name != null && prop.Name == propertyName)
                        {
                            return prop.Value.ToString();
                        }
                    }
                }
            }
            return null;
        }


        //Step through the Word Doc to get the order of the Concepts
        public string GetConceptOrder(Word.Document doc)
        {
            string orderlist = "";
            //Is it one of ours
            if (isTemplate(doc))
            {

                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');

                        if (taga.Length > 1 && taga[0] == "Concept" && taga[1] != "")
                        {
                            orderlist += (orderlist == "" ? "" : ",") + tag;
                        }
                    }
                }
            }
            return orderlist;
        }


        public void SelectContractTemplatesConcept(Word.Document doc,string conceptid)
        {
            //Is it one of ours
            if (isTemplate(doc))
            {
                string templateid = GetDocId(doc);

                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {
                            //Select
                            cc.Range.Select();
                        }
                    }
                }

            }
        }


        public void RemoveConcept(Word.Document doc,string conceptid)
        {
            if (isTemplate(doc))
            {

                //Remove the handler
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                vstoDoc.ContentControlOnEnter -= new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter);
                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                //Create an array of concept content controls in the doc - have to copy or we get into problems with the range including the new ones
                //that we created
                Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
                int cnt = 0;
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    ccs[cnt++] = cc;
                }

                //Now step through all the Contact Controls and update the XML so we get the newest clauses
                foreach (Word.ContentControl cc in ccs)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {

                            

                            //Remove the content control
                            //have to unlock it first
                            cc.LockContents = false;
                            cc.LockContentControl = false;

                            //Also have to unlock any other controls in the range
                            foreach (Word.ContentControl child in cc.Range.ContentControls)
                            {
                                if (child.ID != cc.ID)
                                {
                                    child.LockContents = false;
                                    child.LockContentControl = false;
                                }
                            }
                            
                            cc.Range.Select();
                            Globals.ThisAddIn.Application.Selection.Delete();
                            cc.Delete();

                        }
                    }

                }


                //Add back in the handler
                vstoDoc.ContentControlOnEnter += new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter);
            }

        }


        public void UpdateContractTemplatesConcept(Word.Document doc,string conceptid, string clauseid,string xml,string lastmodified)
        {
            if(doc==null) doc = Application.ActiveDocument;
            //Is it one of ours
            if (isTemplate(doc))
            {

                //Remove the handler
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                try
                {
                    vstoDoc.ContentControlOnEnter -= new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter);
                } catch(Exception){

                }

                string templateid = GetDocId(doc);

                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                //Create an array of concept content controls in the doc - have to copy or we get into problems with the range including the new ones
                //that we created
                Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
                int cnt = 0;
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    ccs[cnt++] = cc;
                }


                //Now step through all the Contact Controls and update the XML so we get the newest clauses
                foreach (Word.ContentControl cc in ccs)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {


                            // check if we have to do this - when the contract is initially loaded the containters are already
                            // populated so don't update them if not required

                            bool selectedclause = false;
                            if (taga.Length > 3)
                            {
                                if (taga[2] == clauseid && lastmodified == taga[3])
                                {
                                    selectedclause = true;
                                }
                            }

                            if (!selectedclause)
                            {

                                //scratch do to hold the clause 
                                Word.Document scratch = Application.Documents.Add(Visible: false);

                                string txt = "";

                                //Get the details of the clause - this would be too chatty when connected to salesforce                            
                                //store the xml in the tree instead and pass into the function

                                //xml = d.GetClauseXML(clauseid).strRtn;

                                if (xml == "")
                                {
                                    xml = "";
                                    txt = "Sorry! problem clause doesn't exist!";
                                }

                                //Populate the content control with the values from the database
                                //have to unlock it first
                                cc.LockContents = false;
                                cc.LockContentControl = false;

                                //Also have to unlock any other controls in the range
                                foreach (Word.ContentControl child in cc.Range.ContentControls)
                                {
                                    if (child.ID != cc.ID)
                                    {
                                        child.LockContents = false;
                                        child.LockContentControl = false;
                                    }
                                }

                                //OK having lots of problems with large paragraphs inserting into the content
                                //control - had a play manually and it worked when cutting and pasting 
                                //*so* get the XML in a sepearate page and then get the formatted text and update
                                //this seems to fix it! need to do some more digging to see if there is a better way
                                Utility.UnlockContentControls(scratch);
                                scratch.Range(scratch.Content.Start, scratch.Content.End).Delete();

                                if (xml != "") scratch.Range().InsertXML(xml);


                                if (xml != "")
                                {
                                    try
                                    {
                                        // delete out what is there
                                        cc.Range.Delete();

                                        // delete out the styles! 
                                        cc.Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

                                        // delete out the pesky tables
                                        for (int tablesi = cc.Range.Tables.Count; tablesi > 0; tablesi--)
                                        {
                                            cc.Range.Tables[tablesi].Delete();
                                        }

                                        // When we insert into the clause it adds a \r - have to get rid of it 
                                        // had a bunch of ways to do it - this seems to work!
                                        Word.Range newr = scratch.Range();
                                        cc.Range.FormattedText = newr.FormattedText;
                                        try{
                                            newr = doc.Range(cc.Range.End-1,cc.Range.End);                                            
                                            if(newr.Characters.Count==1){
                                                if(newr.Characters[1].Text == "\r"){  // Characters starts at 1 - gets me everytime
                                                    newr.Delete();
                                                }
                                            }

                                        } catch(Exception){

                                        }
                                        
                                    }
                                    catch (Exception)
                                    {
                                    }

                                    // close the scratch
                                   var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                                   docclosescratch.Close(false);
                                   System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                                }
                                else
                                {
                                    cc.Range.InsertAfter(txt);
                                }

                                /* DO THIS IN THE SCRATCH NOW! this was causing all sorts of bother
                                //remove any trailing carriage returns
                            
                                for (var i = cc.Range.Characters.Count; i > 0; i--)
                                {
                                    if (i <= cc.Range.Characters.Count)
                                    {
                                        if (cc.Range.Characters[i].Text == "\r")
                                        {
                                            cc.Range.Characters[i].Delete();
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                             
                                /*if (cc.Range.Footnotes.Count == 0)
                                {
                                    while (cc.Range.Characters.Last.Text == "\r")
                                    {
                                        cc.Range.Characters.Last.Delete();
                                    }
                                }
                                */

                                // update the tag
                                cc.Tag = "Concept|" + conceptid + "|" + clauseid + "|" + lastmodified;


                                //relock
                                cc.LockContents = true;
                                cc.LockContentControl = true;

                                
                            }
                        }
                    }

                }

                

                //Add back in the handler
                try
                {
                    vstoDoc.ContentControlOnEnter += new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter);
                }
                catch (Exception)
                {

                }
            }

        }

        public void UpdateContractTemplatesConceptTag(Word.Document doc, string conceptid, string clauseid, string lastmodified)
        {
            if (doc == null) doc = Application.ActiveDocument;
            //Is it one of ours
            if (isTemplate(doc))
            {

               

                string templateid = GetDocId(doc);

                //Now step through the doc and update the concept TAG if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                //Create an array of concept content controls in the doc - have to copy or we get into problems with the range including the new ones
                //that we created
                Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
                int cnt = 0;
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    ccs[cnt++] = cc;
                }

                //Now step through all the Contact Controls and update the XML so we get the newest clauses
                foreach (Word.ContentControl cc in ccs)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {
                            // update the tag
                            if (clauseid == "")
                            {
                                cc.Tag = "Concept|" + conceptid;
                            }
                            else
                            {
                                cc.Tag = "Concept|" + conceptid + "|" + clauseid + "|" + lastmodified;
                            }
                        }
                    }

                }
            }

        }

        public void SelectConcept(string id)
        {
            if (Application.Documents.Count == 0) return;
            Word.Document doc = Application.ActiveDocument;
            //Is it one of ours
            if ((isTemplate(doc) || isContract(doc)))
            {

                //Word.ContentControls cc = doc.SelectContentControlsByTag("Concept|" + id);

                foreach (Word.ContentControl c in doc.Range().ContentControls)
                {
                    if (c.Tag != null)
                    {
                        string tag = Convert.ToString(c.Tag);
                        if (tag.StartsWith("Concept|" + id))
                        {
                            c.Range.Select();
                        }
                    }
                }
            }
        }

        public void SelectElements(string id)
        {
            if (Application.Documents.Count == 0) return;
            Word.Document doc = Application.ActiveDocument;
            //Is it one of ours
            if ((isTemplate(doc) || isClause(doc)))
            {

                //Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Element" && taga[1] == id.ToString())
                        {
                            //Select
                            cc.Range.Select();
                            return;
                        }
                    }
                }

            }
        }

        public void RemoveElements(Word.Document doc,string id)
        {
            //Is it one of ours and is it a clause
            if (isClause(doc) || isTemplate(doc))
            {

                //Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Element" && taga[1] == id.ToString())
                        {
                            //Select
                            cc.Range.Select();

                            cc.LockContents = false;
                            cc.LockContentControl = false;

                            //Also have to unlock any other controls in the range
                            foreach (Word.ContentControl child in cc.Range.ContentControls)
                            {
                                if (child.ID != cc.ID)
                                {
                                    child.LockContents = false;
                                    child.LockContentControl = false;
                                }
                            }

                            cc.Range.Select();
                            Globals.ThisAddIn.Application.Selection.Delete();
                            cc.Delete();
                        }
                    }
                }
            }
        }

        public void MakeDropDownElementsText(Word.Document doc)
        {

            //Comparison doesn't work with drop down controls so get rid of them!

            //Now step through the doc
            object start = doc.Content.Start;
            object end = doc.Content.End;
            Word.Range r = doc.Range(ref start, ref end);

            // Step through and select the one passed
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                string tag = cc.Tag;
                if (tag != null && tag != "" && cc.Tag.Contains('|'))
                {
                    string[] taga = cc.Tag.Split('|');
                    if (taga[0] == "Element")
                    {
                        if (cc.Type == Word.WdContentControlType.wdContentControlDropdownList || cc.Type == Word.WdContentControlType.wdContentControlComboBox)
                        {
                            cc.Type = Word.WdContentControlType.wdContentControlText;
                        }

                    }
                }
            }


        }

        private void UpdateElements(Word.Document doc, Dictionary<string, string> elementValues)
        {

            //Now step through the doc and update the elements
            object start = doc.Content.Start;
            object end = doc.Content.End;
            Word.Range r = doc.Range(ref start, ref end);

            // Step through and select the one passed
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                string tag = cc.Tag;
                if (tag != null && tag != "" && cc.Tag.Contains('|'))
                {
                    string[] taga = cc.Tag.Split('|');
                    if (taga[0] == "Element")
                    {
                        //OK - get the value from the dictionary and format it and update!
                        string id = taga[1];

                        if (elementValues.ContainsKey(id))
                        {
                            string value = elementValues[id];
                            string type = "";

                            try //put in a catch incase there are redline issues!
                            {
                                if (cc.Type == Word.WdContentControlType.wdContentControlComboBox)
                                {
                                    foreach (Word.ContentControlListEntry de in cc.DropdownListEntries)
                                    {
                                        if (de.Text == value) de.Select();
                                    }
                                }
                                else
                                {
                                    //Select
                                    //cc.Range.Select();   
                                    if (value.EndsWith("\\n"))
                                    {
                                        value = value.Substring(0, value.Length - 2) + (char)11;
                                    }

                                    if (value == "") value = " ";

                                    // ok basic formatting support
                                    Word.Font f = cc.Range.Font.Duplicate;
                                    cc.Range.Text = value;
                                    cc.Range.Font = f;
                                    
                                }
                            }
                            catch (Exception e)
                            {
                            }

                        }




                    }
                }
            }


        }

        public void UpdateElement(string id, string value,string type)
        {
            Word.Document doc = Application.ActiveDocument;
            //Is it one of ours
            if (isContract(doc))
            {

                //Now step through the doc and update the elements
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Element" && taga[1] == id.ToString())
                        {
                            //Only do it if there has been a change
                            if (cc.Range.Text != null)
                            {
                                if (cc.Range.Text.Trim() != value.Trim() && !(value.Trim() == "" && cc.Range.Text == cc.Title))
                                {
                                    try //put in a catch incase there are redline issues!
                                    {
                                        if (type == "Picklist" || type == "Checkbox")
                                        {
                                            foreach (Word.ContentControlListEntry de in cc.DropdownListEntries)
                                            {
                                                if (de.Text == value) de.Select();
                                            }
                                        }
                                        else
                                        {
                                            //Select
                                            //cc.Range.Select();   
                                            if (value.EndsWith("\\n"))
                                            {
                                                value = value.Substring(0, value.Length - 2) + (char)11;
                                            }

                                            if (value == "") value = " ";

                                            // ok basic formatting support
                                            Word.Font f = cc.Range.Font.Duplicate;
                                            cc.Range.Text = value;
                                            cc.Range.Font = f;

                                        }
                                    }
                                    catch (Exception)
                                    {
                                    }

                                }
                            }
                        }
                    }
                }

            }
        }

        public string GetElementValue(string id, string type)
        {
            string val = "";
            Word.Document doc = Application.ActiveDocument;
            //Is it one of ours
            if (isContract(doc))
            {

                //Now step through the doc and update the elements
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Element" && taga[1] == id.ToString())
                        {
                            val = cc.Range.Text;
                        }
                    }
                }

            }
            return val;
        }
        public void InitiateElement(string id, string value,string type,string format,string[] options,string option1,string option2)
        {
            //Find each element and update the Content Control to the right thing

            Word.Document doc = Application.ActiveDocument;
            //Is it one of ours
            if (isContract(doc))
            {

                //Now step through the doc and update the elements
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Element" && taga[1] == id.ToString())
                        {



                            if (type == "Picklist")
                            {
                                if (cc.Type != Word.WdContentControlType.wdContentControlDropdownList) cc.Type = Word.WdContentControlType.wdContentControlDropdownList;
                                cc.DropdownListEntries.Clear();
                                foreach (string entry in options)
                                {
                                    if(entry!=""){
                                    bool alreadythere = false;
                                    foreach (Word.ContentControlListEntry de in cc.DropdownListEntries)
                                    {
                                        if (de.Text == entry) alreadythere = true;
                                    }
                                    if(!alreadythere)cc.DropdownListEntries.Add(entry,entry);
                                    }
                                }
                                cc.LockContents = false;
                                cc.LockContentControl = true;

                                //set the value
                                if (cc.Range.Text==null || cc.Range.Text.Trim() != value)
                                {
                                    foreach (Word.ContentControlListEntry de in cc.DropdownListEntries)
                                    {
                                        if (de.Text == value) de.Select();
                                    }
                                }
                            }
                            else if (type == "Date")
                            {
                                if (cc.Type != Word.WdContentControlType.wdContentControlDate) cc.Type = Word.WdContentControlType.wdContentControlDate;
                                cc.DateDisplayFormat = format;
                                cc.LockContents = false;
                                cc.LockContentControl = true;

                                if (cc.Range.Text.Trim() != value)
                                {
                                    // ok basic formatting support
                                    Word.Font f = cc.Range.Font.Duplicate;
                                    cc.Range.Text = value;
                                    cc.Range.Font = f;
                                }
                            }
                            else if (type == "Checkbox")
                            {
                                if (cc.Type != Word.WdContentControlType.wdContentControlDropdownList) cc.Type = Word.WdContentControlType.wdContentControlDropdownList;
                                cc.DropdownListEntries.Clear();
                                if (option1 != "")
                                {
                                    cc.DropdownListEntries.Add(option1, "Checked");
                                    if (option1 != option2)
                                    {
                                        if (option2 == "") option2 = " ";
                                        cc.DropdownListEntries.Add(option2, "UnChecked");
                                    }
                                }
                                cc.LockContents = false;
                                cc.LockContentControl = true;

                                //set the value
                                if (cc.Range.Text.Trim() != value)
                                {
                                    foreach (Word.ContentControlListEntry de in cc.DropdownListEntries)
                                    {
                                        if (de.Text == value) de.Select();
                                    }
                                }
                            }
                            else
                            {
                                // only change this if we need to
                                if (cc.Type!=Word.WdContentControlType.wdContentControlText) cc.Type = Word.WdContentControlType.wdContentControlText;
                                cc.LockContents = false;
                                cc.LockContentControl = true;

                                //also set the value
                                if (value.EndsWith("\\n"))
                                {
                                    value = value.Substring(0, value.Length - 2) + (char)11;
                                }
                                if (cc.Range.Text != null && cc.Range.Text.Trim() != value)
                                {
                                    // if the value is blank then put in a space so we don't loose formatting
                                    if (value == "") value = " ";

                                    // ok basic formatting support
                                   
                                    Word.Font f = cc.Range.Font.Duplicate;
                                    cc.Range.Text = value;
                                    cc.Range.Font = f;


                                }
                            }
                        }
                    }
                }

            }
        }
     

        //Contract Instance Stuff - might be the same as ContractTemplate
        public void UpdateContractConcept(string conceptid, string clauseid, string xml,string lastmodified, Word.Document doc, Dictionary<string, string> elementValues)
        {
            //Is it one of ours
            if (isContract(doc))
            {
                //Remove the handler
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                vstoDoc.ContentControlOnExit -= new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                string contractid = GetDocId(doc);

                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);

                //Create an array of concept content controls in the doc - have to copy or we get into problems with the range including the new ones
                //that we created
                Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
                int cnt = 0;
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    ccs[cnt++] = cc;
                }

                Globals.ThisAddIn.Application.ScreenUpdating = false;

                //scratch do to hold the clause 
                Word.Document scratch = Application.Documents.Add(Visible: false);

                //hold the old clauses xml;
                string oldxml = "";

                //Now step through all the Contact Controls and update the XML so we get the newest clauses
                foreach (Word.ContentControl cc in ccs)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {

                            // Get the details of the clause - this would be too chatty when connected to salesforce                            
                            // store the xml in the tree instead and pass into the function                            
                            // DataReturn dr = _d.GetClause(clauseid);

                            string txt = "";
                            if (clauseid!="" && xml == "")
                            {
                                xml = "";
                                txt = "Sorry! problem clause doesn't exist!";
                            }

                            if (cc.PlaceholderText != null)
                            {
                                cc.SetPlaceholderText(Text: "");
                            }

                            // Populate the content control with the values from the database
                            // have to unlock it first
                            cc.LockContents = false;
                            cc.LockContentControl = true;

                            //Also have to unlock any other controls in the range
                            foreach (Word.ContentControl child in cc.Range.ContentControls)
                            {
                                if (child.ID != cc.ID)
                                {
                                    child.LockContents = false;
                                    child.LockContentControl = false;
                                }
                            }

                            oldxml = cc.Range.WordOpenXML;

                            // OK having lots of problems with large paragraphs inserting into the content
                            // control - had a play manually and it worked when cutting and pasting 
                            // *so* get the XML in a sepearate page and then get the formatted text and update
                            // this seems to fix it! need to do some more digging to see if there is a better way

                            Utility.UnlockContentControls(scratch);
                            scratch.Range(scratch.Content.Start, scratch.Content.End).Delete();
                            if (xml != "") scratch.Range(0).InsertXML(xml);

                            // if clauseid is blank then its a "select none" so do it even though the xml is blank
                            if (clauseid=="" || xml != "")
                            {
                                //cc.Range.InsertXML(xml);
                                
                                //Track changes?
                                if (doc.TrackRevisions)
                                {

                                    // OK - gets more complicated! save the old paragraph to a scratch doc and undo any changes
                                    // then get the new one in another scratch doc - stop tracking changes in the current doc
                                    // do a diff and then insert that in the parra
                                    
                                    Word.Document oldclause=Globals.ThisAddIn.Application.Documents.Add(Visible: false);            
                                    string oldclausefilename = Utility.SaveTempFile(doc.Name + "-oldclause");
                                    oldclause.Range().InsertXML(oldxml);
                                    
                                    // get rid of any changes - have to make it the active doc to do this
                                    oldclause.Activate();
                                    oldclause.RejectAllRevisions();

                                    MakeDropDownElementsText(oldclause);

                                    // Now update the elements of the scratch
                                    Utility.UnlockContentControls(scratch);
                                    UpdateElements(scratch, elementValues);
                                    // Dropdowns don't diff well (they show as changes, so change the content controls to text - they'll get changed back by initiate)
                                    MakeDropDownElementsText(scratch);

                                    // Now run a diff - do it from the old doc rather than a compare so it gives us the redline rather than blue line compare
                                    string scratchfilename = Utility.SaveTempFile(doc.Name + "-newclause");
                                    scratch.SaveAs2(FileName: scratchfilename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);                                    
                                    // this is how you do it as a pure compare - Word.Document compare = Application.CompareDocuments(oldclause, scratch,Granularity:Word.WdGranularity.wdGranularityCharLevel);
                                    
                                    oldclause.Compare(scratchfilename, CompareTarget: Word.WdCompareTarget.wdCompareTargetCurrent, AddToRecentFiles: false);                                    
                                    oldclause.ActiveWindow.Visible = false;

                                    // Activate the doc - switch of tracking and insert the marked up dif
                                    doc.Activate();
                                    doc.TrackRevisions = false;

                                    // delete out what is there
                                    cc.Range.Delete();

                                    // delete out the styles! 
                                    cc.Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

                                    // delete out the pesky tables
                                    for (int tablesi = cc.Range.Tables.Count; tablesi > 0; tablesi--)
                                    {
                                        cc.Range.Tables[tablesi].Delete();
                                    }

                                    cc.Range.FormattedText = oldclause.Content.FormattedText;
                                    doc.Activate();
                                    doc.TrackRevisions = true;

                                    var doccloseoldclause = (Microsoft.Office.Interop.Word._Document)oldclause;
                                    doccloseoldclause.Close(false);


                                }
                                else
                                {
                                    try
                                    {
                                        // delete out what is there
                                        cc.Range.Delete();

                                        // delete out the styles! 
                                        cc.Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

                                        // delete out the pesky tables
                                        for (int tablesi = cc.Range.Tables.Count; tablesi > 0; tablesi--)
                                        {
                                            cc.Range.Tables[tablesi].Delete();
                                        }

                                        Word.Range newr = scratch.Range();
                                        cc.Range.FormattedText = newr.FormattedText;
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                            }
                            else
                            {
                                cc.Range.InsertAfter(txt);
                            }

                            // sort out the formatting problems caused by inserting into the container                           
                            doc.Activate();
                            bool tchanges = doc.TrackRevisions;
                            doc.TrackRevisions = false;


                            // When we insert into the clause it adds a \r - have to get rid of it 
                            // had a bunch of ways to do it - this seems to work!
                            try
                            {
                                Word.Range newr = doc.Range(cc.Range.End - 1, cc.Range.End);
                                if (newr.Characters.Count == 1)
                                {
                                    if (newr.Characters[1].Text == "\r")
                                    {  // Characters starts at 1 - gets me everytime
                                       newr.Delete();
                                    }
                                }

                            }
                            catch (Exception)
                            {

                            }

                            doc.TrackRevisions = tchanges;
                            

                            //close the scratch
                            var docclosescratch = (Microsoft.Office.Interop.Word._Document)scratch;
                            docclosescratch.Close(false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(docclosescratch);

                            // update the tag
                            cc.Tag = "Concept|" + conceptid + "|" + clauseid + "|" + lastmodified;

                            //relock
                            cc.LockContents = true;
                            cc.LockContentControl = true;
                        }
                    }

                }

                Globals.ThisAddIn.Application.ScreenUpdating = true;

                //Add back in the handler
                vstoDoc.ContentControlOnExit += new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

            }

        }


        public string GetContractClauseText(Word.Document doc,string conceptid)
        {
            return GetContractClauseRange(doc,conceptid).Text;
        }

        public string GetContractClauseXML(Word.Document doc,string conceptid)
        {
            return GetContractClauseRange(doc,conceptid).WordOpenXML;
        }

        public Word.Range GetContractClauseRange(Word.Document doc,string conceptid)
        {
            Word.Range r = null;
            //Is it one of ours
            if (isContract(doc))
            {

                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {
                            //Get Text
                            r = cc.Range;
                        }
                    }
                }

            }

            return r;
        }


        public string GetTemplateClauseText(Word.Document doc, string conceptid)
        {
            Word.Range r = GetTemplateClauseRange(doc, conceptid);
            if (r != null)
            {
                return r.Text;
            }
            else
            {
                return "";
            }
        }

        public string GetTemplateClauseXML(Word.Document doc, string conceptid)
        {
            Word.Range r = GetTemplateClauseRange(doc,conceptid);
            if (r != null)
            {
                return r.WordOpenXML;
            }
            else
            {
                return "";
            }
        }

        public Word.Range GetTemplateClauseRange(Word.Document doc, string conceptid)
        {

            Word.Range r = null;
            //Is it one of ours
            if (isTemplate(doc))
            {

                //Now step through the doc and update the concept if it matches the one we just updated
                object start = doc.Content.Start;
                object end = doc.Content.End;
                r = doc.Range(ref start, ref end);

                // Step through and select the one passed
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {
                            //Get Text
                            r = cc.Range;
                        }
                    }
                }

            }

            return r;
        }

        public void UnlockContractConcept(string conceptid,Word.Document doc)
        {
            //Is it one of ours
            if (isContract(doc))
            {
                //Remove the handler
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                vstoDoc.ContentControlOnExit -= new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                string contractid = GetDocId(doc);

                //Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);


                Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];               
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {

                            // unlock the clause
                            cc.LockContents = false;
                            cc.LockContentControl = true;

                            // also have to unlock any other controls in the range
                            foreach (Word.ContentControl child in cc.Range.ContentControls)
                            {
                                if (child.ID != cc.ID)
                                {
                                    child.LockContents = false;
                                    child.LockContentControl = true;
                                }
                            }

                            // update the tag to set the modified to be unlocked so we know to load the clause
                            // from the database
                            cc.Tag = "Concept|" + taga[1].ToString() + "|" + taga[2].ToString() + "|" + "Unlocked";

                        }

                    }
                    //Add back in the handler
                    vstoDoc.ContentControlOnExit += new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                }
            }
          

        }

        public void UnlockLockTemplateConcept(Word.Document doc,string conceptid,bool lck)
        {
            //Is it one of ours
            if (isTemplate(doc))
            {
                //Remove the handler
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                vstoDoc.ContentControlOnExit -= new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                string templateid = GetDocId(doc);

                //Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);


                Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
                int cnt = 0;
                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');
                        if (taga[0] == "Concept" && Convert.ToString(taga[1]) == conceptid)
                        {

                            cc.LockContents = lck;
                            cc.LockContentControl = true;

                        }
                    }
                    //Add back in the handler
                    vstoDoc.ContentControlOnExit += new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                }
            }
        }

        //--- Intercept save handler

        void Application_DocumentBeforeSave(Word.Document doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //Check if this is one of our docs and if it is then do the right thing, not the black and white thing
            //Add a document property so we know that is a contract template and what the id is
            try
            {
                bool hidep = false;
                if (!Globals.ThisAddIn._p.IsVisible)
                {
                    Globals.ThisAddIn.ProcessingStart("Saving");
                    hidep = true;
                }

                string prop = GetCurrentAxiomDocProp();
                if (prop != null)
                {
                    string[] propa = prop.Split('|');
                    if (propa[0] == "ContractTemplate")
                    {
                        Globals.ThisAddIn.ProcessingUpdate("Save Contract Template");

                        // Get the Sidebar and save the elemnt value if that has changed
                        TemplateEdit.TEditSidebar tsb = Globals.ThisAddIn.GetTaskPaneControlTemplate(doc);
                        if(tsb!=null) tsb.FormSave();

                        //save this to a scratch file
                        Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                        string filename = Utility.SaveTempFile(propa[1]);
                        doc.SaveAs2(FileName: filename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                        //Save a copy!
                        Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                        string filenamecopy = Utility.SaveTempFile(propa[1] + "X");
                        Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: false);
                        dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                        var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                        docclose.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                        //Now 
                        Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                        _d.SaveTemplateFile(propa[1], filenamecopy);

                        //d.SaveTemplateXML(propa[1], doc.WordOpenXML);


                        //Cancel the save
                        SaveAsUI = false;
                        Cancel = true;
                    }

                    if (propa[0] == "ClauseTemplate")
                    {
                        //Save!
                        //doc = Globals.ThisAddIn.Application.ActiveDocument;
                        //d.SaveClauseXML(propa[1],doc.Content.Text ,doc.WordOpenXML);
                        Globals.ThisAddIn.ProcessingUpdate("Save Clause Template");
                        // doc = Globals.ThisAddIn.Application.ActiveDocument;

                        //save this to a scratch file
                        Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                        string filename = Utility.SaveTempFile(propa[1]);
                        doc.SaveAs2(FileName: filename, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                        //Save a copy!
                        Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                        string filenamecopy = Utility.SaveTempFile(propa[1] + "X");
                        Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(filename, Visible: false);
                        dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                        var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                        docclose.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                        //Now save the file and then refresh any other docsto reflect the change
                        Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                        _d.SaveClauseFile(propa[1], doc.Content.Text, filenamecopy);

                        Globals.ThisAddIn.ProcessingUpdate("Update Any Templates with the clause");
                        RefreshAllTaskPanesWithClause(propa[1], doc.WordOpenXML);
                        doc.Activate();

                        //Cancel the save
                        SaveAsUI = false;
                        Cancel = true;
                    }

                    if (propa[0] == "Contract" || propa[0] == "UAContract")
                    {
                        //Save the doc and the data
                        GetTaskPaneControlContract().SaveContract(false,true);

                        //Cancel the save
                        SaveAsUI = false;
                        Cancel = true;
                    }

                    if (propa[0] == "Compare")
                    {
                        //Save the doc and the data
                        GetTaskPaneControlCompare().SaveContract(false, true);

                        //Cancel the save
                        SaveAsUI = false;
                        Cancel = true;
                    }


                    if (hidep)
                    {
                        Globals.ThisAddIn.ProcessingStop("Stop");
                    }

                }
            }
            catch (Exception e)
            {
                Globals.ThisAddIn.ProcessingStop("Stop");
                System.Windows.MessageBox.Show("Sorry there has been a problem:"+e.Message);

            }
        }





        public void SetTheme()
        {
            if ( _localSettings.Theme == LocalSettings.Themes.Windows8)
            {
                StyleManager.ApplicationTheme = new Windows8Theme();
                //Set default Theme            
                Color myRgbColor = new Color();

                string color = _localSettings.ThemeColor;
                if (color == "") color = "#DE5827";
                try
                {
                    myRgbColor = (Color)ColorConverter.ConvertFromString(color);
                    Windows8Palette.Palette.AccentColor = myRgbColor;
                }
                catch (Exception)
                {

                }
            }
            else if (_localSettings.Theme == LocalSettings.Themes.Dark)
            {
                StyleManager.ApplicationTheme = new Expression_DarkTheme();
            }
            else if (_localSettings.Theme == LocalSettings.Themes.Office)
            {
                StyleManager.ApplicationTheme = new Office_BlackTheme();
            }
            
        }

        public void OpenReports()
        {
            LocalSettings.Instances ?Inst = _localSettings.Inst;
            string reportsurl = "";


            // Get the reports url from the settings, if not there default to the prod reports
            JToken s = Globals.ThisAddIn.GetSettings().GetGeneralSetting("Reports");
            if (s == null)
            {
                reportsurl = "https://reports.irisbyaxiom.com/SSOLogin.aspx";
            }
            else
            {
                reportsurl = s.ToString();
            }

            reportsurl += "?sfid=" + _d.GetSessionId() + "&sfurl=" + Uri.EscapeUriString(_d.GetPartnerURL());
            System.Diagnostics.Process.Start(reportsurl);
        }


        public void RemoveContentControls(Word.Document doc)
        {
            //Is it one of ours
            if (isContract(doc))
            {
                //Remove the handler
                Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                vstoDoc.ContentControlOnExit -= new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                string contractid = GetDocId(doc);

                //Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);


                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');

                        if (taga.Length > 1 && ((taga[0] == "Concept" && taga[1] != "") || (taga[0] == "Element" && taga[1] != "")))
                        {
                            Word.Range ccr = cc.Range;
                            cc.LockContentControl = false;
                            cc.LockContents = false;                            
                            cc.Delete(false);

                            // *TODO* if none selected we may want to delete the extra return
                        }
                    }
                }
            }
        }





        // this is used in the Template Import - updates the ContentControl tags to match the new
        // ids that have been created - this one does both the Clauses and the Elements that are 
        // being displayed in the current template clause selection

        public void UpdateContractTemplateContentControls(Word.Document doc, Dictionary<string, string> ConceptMapping,Dictionary<string, string> ClauseMapping, Dictionary<string, string> ElementMapping)
        {

            //Is it one of ours
            if (isTemplate(doc))
            {
                //Remove the handler - won't have any hanlders yet - loaded it as an exporttemplate
                // Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                // vstoDoc.ContentControlOnExit -= new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                string contractid = GetDocId(doc);

                //Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);


                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');

                        if (taga.Length > 1 && ((taga[0] == "Concept" && taga[1] != "")))
                        {
                            // Concept is the format Concept|ConceptId|SelectedClauseId|LastModified
                            // as we are importing this everything should be up to date - set the LastModified to "0000"
                            // to indicate it is up to date - need to do this cause I don't have the last modified date

                            string ConceptIdOld = taga[1].ToString();
                            string ClauseIdOld = "";
                            if (taga.Length > 2)
                            {
                                ClauseIdOld = taga[2].ToString();
                            }

                            // if we have the new clause id then set the tag so that we know its the one in the doc
                            // "0000" indicates it has been loaded so is up to date
                            if (ClauseMapping.ContainsKey(ClauseIdOld))
                            {
                                cc.Tag = "Concept|" + ConceptMapping[ConceptIdOld] + "|" + ClauseMapping[ClauseIdOld] + "|0000";
                            }
                            else
                            {
                                cc.Tag = "Concept|" + ConceptMapping[ConceptIdOld];
                            }

                        }
                        else if ((taga[0] == "Element" && taga[1] != ""))
                        {
                            // Element tag is the format Element|ElementId|ClauseId                           
                            string ElementIdOld = taga[1].ToString();
                            string ClauseIdOld = taga[2].ToString();

                            // if its not in the mapping leave it be
                            if (ElementMapping.ContainsKey(ElementIdOld))
                            {
                                if (ClauseMapping.ContainsKey(ClauseIdOld))
                                {
                                    cc.Tag = "Element|" + ElementMapping[ElementIdOld] + "|" + ClauseMapping[ClauseIdOld];
                                }
                                else
                                {
                                    cc.Tag = "Element|" + ElementMapping[ElementIdOld];
                                }
                            }
                        }
                    }
                }
            }
        }

        public void UpdateClauseTemplateContentControls(Word.Document doc, Dictionary<string, string> ConceptMapping, Dictionary<string, string> ClauseMapping, Dictionary<string, string> ElementMapping)
        {

            //Is it one of ours
            if (isClause(doc))
            {
                // Remove the handler - won't have any hanlders yet - loaded it as an exporttemplate
                // Microsoft.Office.Tools.Word.Document vstoDoc = Globals.Factory.GetVstoObject(doc);
                // vstoDoc.ContentControlOnExit -= new Word.DocumentEvents2_ContentControlOnExitEventHandler(vstoDoc_ContentControlOnExit);

                // Now step through the doc
                object start = doc.Content.Start;
                object end = doc.Content.End;
                Word.Range r = doc.Range(ref start, ref end);


                foreach (Word.ContentControl cc in r.ContentControls)
                {
                    string tag = cc.Tag;
                    if (tag != null && tag != "" && cc.Tag.Contains('|'))
                    {
                        string[] taga = cc.Tag.Split('|');

                        if ((taga[0] == "Element" && taga[1] != ""))
                        {
                            // Element tag is the format Element|ElementId|ClauseId

                            string ElementIdOld = taga[1].ToString();
                            string ClauseIdOld = taga[2].ToString();

                            // if its not in the mapping leave it be
                            if (ElementMapping.ContainsKey(ElementIdOld))
                            {
                                if (ClauseMapping.ContainsKey(ClauseIdOld))
                                {
                                    cc.Tag = "Element|" + ElementMapping[ElementIdOld] + "|" + ClauseMapping[ClauseIdOld];
                                }
                                else
                                {
                                    cc.Tag = "Element|" + ElementMapping[ElementIdOld];
                                }
                            }
                        }
                    }
                }
            }
        }


        public string GetFootnotes(Word.Document doc,string ConceptId)
        {

            string rtn = "";
            if (isTemplate(doc))
            {

                foreach (Word.ContentControl c in doc.Range().ContentControls)
                {
                    if (c.Tag != null)
                    {
                        string tag = Convert.ToString(c.Tag);
                        if (tag.StartsWith("Concept|" + ConceptId))
                        {
                            if (c.Range.Footnotes.Count > 0)
                            {
                                for (int z = 1; z <= c.Range.Footnotes.Count; z++)
                                {
                                    string foot = c.Range.Footnotes[z].Range.Text;
                                    if(foot.StartsWith("\t")) foot = foot.Substring(1);
                                    rtn += foot + "\n";
                                }
                            }


                        }
                    }
                }
            }
            else if (isClause(doc))
            {
                for (int z = 1; z <= doc.Range().Footnotes.Count; z++)
                {
                    string foot = doc.Range().Footnotes[z].Range.Text;
                    if (foot.StartsWith("\t")) foot = foot.Substring(1);
                    rtn += foot + "\n";                    
                }
            }

            return rtn;
        }

        public void ScreenUpdatingOff()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
        }

        public void ScreenUpdatingOn()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }


        public void OpenEditWindow(string Name){
            CloseEditWindows();
            if (_editWindows.ContainsKey(Name))
            {                
                Edit ewin = _editWindows[Name];
                ewin.Show();
                ewin.Focus();
            }
            else
            {
                Edit ewin = new Edit(Name);
                _editWindows.Add(Name, ewin);
                ewin.Show();
                ewin.Focus();
            }
        }

        public void CloseEditWindows()
        {
            foreach (string key in _editWindows.Keys)
            {
                Edit ewin = _editWindows[key];
                ewin.Close();
            }
        }

        public void OpenZoomEditWindow(string Name,string Id)
        {
            CloseEditZoomWindows(Name);
            if (_editZoomWindows.ContainsKey(Name))
            {
                Edit ewin = _editZoomWindows[Name];
                ewin.OpenZoomEditId(Id);
            }
            else
            {
                // make sure we have the object set up                
                string[] sObj = Globals.ThisAddIn.GetSettings().GetSetting("", "SObjects").Split('|');
                if (sObj.Contains(Name))
                {
                    Edit ewin = new Edit("Zoom", Name, Id);
                    _editZoomWindows.Add(Name, ewin);
                    
                }
                else
                {
                    // open as a webpage
                    // need to get the url
                    string url = _d.GetUrlForNonLoaded(Name);
                    Uri temp = new Uri(url.Replace("{ID}", Id));
                    string rooturl = temp.Scheme + "://" + temp.Host;

                    string frontdoor = rooturl + "/secur/frontdoor.jsp?sid=" + _d.GetSessionId();
                    string redirect = frontdoor + "&retURL=" + temp.PathAndQuery;

                    System.Diagnostics.Process.Start(redirect);
                }
            }
        }

        public void CloseEditZoomWindows()
        {
            foreach (string key in _editZoomWindows.Keys)
            {
                Edit ewin = _editZoomWindows[key];
                ewin.Close();
            }
        }

        public void CloseEditZoomWindows(string except)
        {
            foreach (string key in _editZoomWindows.Keys)
            {
                if (key != except)
                {
                    Edit ewin = _editZoomWindows[key];
                    ewin.Close();
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Logger.Init();
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
