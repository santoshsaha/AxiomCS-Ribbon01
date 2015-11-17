using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AxiomIRISRibbon.sfPartner;

namespace AxiomIRISRibbon
{

    public class DataReturn
    {
        public DataTable dt;
        public string strRtn;
        public int intRtn;
        public string id;
        public bool success;
        public string reload;
        public bool fromcache;
        public string errormessage;

        public DataReturn()
        {
            dt = new DataTable();
            strRtn = "";
            success = true;
            fromcache = false;
            reload = "";
            errormessage = "";
            id = "";
        }
    }


    public class Data
    {
        // This class acts as a frontend to the DataStore - currently only supports Salesforce 
        // but all calls go through here to Salesforce, so could support diferent data stores 
        // by adding switches here - the Salesforce edit front end does rely heavilly on the
        // Salesforce API though so would be tricky!

        // Cache - just cache each object and invalidate on a change
        // Actually would be better to hold this as local data in a DataSet
        // but then we would have to know the datamodel

        // Russel 18 October, 2015 - this used to (sort) of support an Access database as well
        // as Salesforce - originally we could use the Access database to demo without a 
        // remote connection but it only worked for the ClauseLibrary bit - Tidy up by removing        

        private Dictionary<string, DataTable> _cache;

        public SalesForce _sf;

        // ok - define the table names as strings - we had instances with inconsitant naming
        // so this let us support diferent names in diferent instances - this may be redundant
        // now - could clean up later
        public string ribbontemplate;
        public string ribbonconcept;
        public string ribbonclause;
        public string ribbonelement;
        public string ribbontemplateconcept;
        public string ribbontemplateclause;
        public string ribbonclauseelement;
        public string version;
        public string clause;
        public string element;
        public string Matter;

        public string contractfilename = "TemplateDocument.docx";
        public string templatefilename = "Template.docx";

        public string demoinstance = "";


        public Data()
        {
            _sf = new SalesForce();
            _cache = new Dictionary<string, DataTable>();
        }

        public void CheckTableNames()
        {
            // these were the original table names - was a bit inconsistant with the plural
            this.ribbontemplate = "RibbonTemplate__c";
            this.ribbonconcept = "RibbonConcepts__c";
            this.ribbonclause = "RibbonClauses__c";
            this.ribbonelement = "RibbonElement__c";
            this.ribbontemplateconcept = "RibbonTemplateConcept__c";
            this.ribbontemplateclause = "RibbonTemplateClause__c";
            this.ribbonclauseelement = "RibbonClauseElement__c";
            this.version = "Version__c";
            this.clause = "Clause__c";
            this.element = "Element__c";
            this.Matter = "Matter__c";

            //check if the tables aren't there, then drop the plural and manage the document versus version thing
            if (!_sf._allSObjects.ContainsKey(ribbonconcept))
            {
                this.ribbonconcept = "RibbonConcept__c";
            }

            if (!_sf._allSObjects.ContainsKey(ribbonclause))
            {
                this.ribbonclause = "RibbonClause__c";
            }

            // hacks to get things to work with the demo instances
            JToken s = Globals.ThisAddIn.GetSettings().GetGeneralSetting("Demo");
            if (s != null)
            {
                if (s.ToString().ToLower() == "general")
                {
                    this.demoinstance = "general";
                    this.version = "Document__c";
                }
                else if (s.ToString().ToLower() == "isda")
                {
                    this.demoinstance = "isda";
                    this.version = "Document__c";
                }
            }
        }

        public string Login(string Username, string Password, string Token, string Url, string InstanceDesc)
        {
            return _sf.Login(Username, Password, Token, Url, InstanceDesc);
        }

        public string Login(string Token, string Url, string MetaUrl, string InstanceDesc)
        {
            return _sf.Login(Token, Url, MetaUrl, InstanceDesc);
        }

        public void Logout()
        {
            //Logout
            _sf.Logout();
        }

        public string GetUser()
        {
            return _sf.GetUser();
        }

        public string GetUserId()
        {
            return _sf.GetUserId();
        }

        public string GetUserProfile()
        {
            return _sf.GetUserProfile();
        }

        public DataReturn GetPickListValues(string sObject, string fName, bool reload)
        {
            if (!reload && _cache.ContainsKey("PickList|" + sObject + "|" + fName))
            {
                DataReturn dr = new DataReturn();
                dr.dt = _cache["PickList|" + sObject + "|" + fName];
                dr.success = true;
                dr.fromcache = true;
                return dr;
            }
            else
            {
                DataReturn dr = _sf.GetPickListValues(sObject, fName);
                if (dr.success) _cache["PickList|" + sObject + "|" + fName] = dr.dt;
                return dr;
            }
        }
        //Code PES
        public DataReturn GetAgreementsForVersion(string id)
        {
            //  return _sf.RunSOQL("SELECT  version_number__c,Id,Template__c FROM version__c WHERE  matter__c='" + id + "' and version_number__c !=null order by version_number__c desc limit 1");
            return _sf.RunSOQL("SELECT  matter__c,Name,Additional_Notes__c,Agreement_Number__c,Applicable_Change_of_Control__c,Assigned_ATM__c,Assignment_Addressed__c,Assignment_Express_Consent_Required__c,Assignment_Notice_Lead_Time_days__c,Assignment_Written_Notice_Required__c,Auto_Renewal_Notification_Days__c,Auto_Renewal_Option__c,Auto_Renewal_Term_Months__c,Auto_Renewal_Terms__c,Average_Annual_Contract_Value__c,Breach_Notice_Requirement__c,Breach_Termination_Notice_Days__c,Breach_Termination_Notice_Required__c,Change_Control_Express_Consent_Lead_Time__c,Conditions_to_Exception__c,Consent_Unreasonably_Withheld__c,Consequences_of_Violating_Assignment_Pro__c,Contract_End_Date__c,Contract_Start_Date_sec1__c,Convenience_Notice_Requirement__c,Convenience_Termination_Notice_Days__c,Convenience_Termination_Notice_Required__c,Date_Terminated__c,Exception_to_Consent__c,Exclusivity_Language__c,Expiration_Notification_Days__c,Explain_Exclusivity_Language__c,Explain_Non_Compete_Language__c,Explain_Terminated_for_Other__c,Explain_Termination_Right_Trggers__c,Express_Consent_Required__c,Express_Written_Contract_Lag_Time_Days__c,External_ID__c,Maximum_Contract_Value_if_capped__c,Non_Compete_Language__c,Non_Standard_Language__c,Num_Days_Notice_to_Initiate_Manual_Renew__c,Number_of_Days_Notice_to_Stop_Auto_Renew__c,Off_Playbook_Language__c,Other_Restrictions_Assign_or_Asset_Tran__c,Perpetual__c,Renewal_Type__c,Requires_Express_Consent__c,Require__c,Template__c,Risk_Rating__c,Status__c,Term_Months__c,Terminated_For__c,Termination_Comments__c,Termination_for_Breach_Option__c,Termination_for_Convenience_Option__c,Termination_Notes__c,Termination_Notice_Days__c,Termination_Notice_Issue_Datec__c,Termination_Notice_Required__c,Termination_Notice_Requirement__c,Termination_Option__c,Termination_Right_Triggers__c,Total_Contract_Value__c,TS_of_Axiom_Ack_to_Receipt_of_CC_s__c,TS_of_CC_s_Received_on_Draft__c,version_number__c,Written_Notice_Lead_Time_days__c,Id FROM version__c WHERE  matter__c='" + id + "' and version_number__c !=null order by version_number__c desc limit 1");
        
        }

        public DataReturn GetAllAttachments(string VersionNumber)
        {
            
            String soql = "select Id, Name, Body, ContentType from Attachment where parentId = '" + VersionNumber + "'";
            return _sf.RunSOQL(soql) ;

         
        }

        public DataReturn GetTemplateAttach(string TemplateId)
        {

            String soql = "select Id, Name, Body, ContentType from Attachment where parentId = '" + TemplateId + "'";
            return _sf.RunSOQL(soql);
        }
        public DataReturn GetAgreementType(string MatterId)
        {

            DataReturn dr = _sf.RunSOQL("select Id,Name,Master_Agreement_Type__c from Matter__c where id = '" + MatterId + "' and IsDeleted = false");
            return dr;
        }
        //End Code PES
        public DataReturn GetTemplates(bool published)
        {
            return _sf.RunSOQL("SELECT Id,Name,Description__c,Type__c,State__c,PlaybookLink__c FROM " + this.ribbontemplate + (published ? " where State__c='Published'" : "") + " order by Name ");
        }
        //Code PES
        public DataReturn GetTemplatesFromExsisting(bool published)
        {
            return _sf.RunSOQL("SELECT  Id,Name,Counterparty__c,Credit_Suisse_Entity__c,CNID__c,Agreement_Number__c  FROM " + this.Matter + " order by Name ");
        }

        //public DataReturn GetTemplateForsearch(string agreementnumber, string cnid)
        //{
        //    string query = "SELECT Name,Counterparty__c,Credit_Suisse_Entity__c,CNID__c,Agreement_Number__c  FROM Matter__c where  CNID__c= '" + cnid + "'";
        //    return _sf.RunSOQL(query);
        //}

        public DataReturn GetTemplateForsearch(string agreementnumber, string CNID)
        {
            string query = string.Empty;
            if (String.IsNullOrEmpty(CNID) && !String.IsNullOrEmpty(agreementnumber))
            {
                query = "SELECT Id,Name,Counterparty__c,Credit_Suisse_Entity__c,CNID__c,Agreement_Number__c  FROM " + this.Matter + " where Agreement_Number__c like " + "'%" + agreementnumber + "%'";
            }
            else if (String.IsNullOrEmpty(agreementnumber) && !String.IsNullOrEmpty(CNID))
            {
                query = "SELECT  Id,Name,Counterparty__c,Credit_Suisse_Entity__c,CNID__c,Agreement_Number__c  FROM " + this.Matter + " where CNID__c like " + "'%" + CNID + "%'";
            }
            else
                query = "SELECT  Id,Name,Counterparty__c,Credit_Suisse_Entity__c,CNID__c,Agreement_Number__c  FROM " + this.Matter + " where Agreement_Number__c like " + "'%" + agreementnumber + "%'" + " AND CNID__c like " + "'%" + CNID + "%'";
            return _sf.RunSOQL(query);

        }

        //End Code PES
        public DataReturn GetTemplate(string Id)
        {
            return _sf.RunSOQL("SELECT Id,Name,Description__c,Type__c,State__c,PlaybookLink__c FROM " + this.ribbontemplate + " where Id = '" + Id + "'");
        }


        public DataReturn CheckTemplate(string name)
        {
            name = Utility.FixUpSOQLString(name.Trim());
            return _sf.RunSOQL("SELECT Id from " + this.ribbontemplate + " where Name = '" + name + "'");
        }


        public DataReturn SaveTemplate(DataRow dr)
        {
            return _sf.Save("RibbonTemplate__c", dr);
        }

        public DataReturn SaveTemplateFile(string Id, string filename)
        {
            return _sf.SaveAttachmentFile(Id, templatefilename, filename);
        }


        public DataReturn DeleteTemplate(String Id)
        {
            return _sf.Delete("RibbonTemplate__c", Id);
        }


        public DataReturn GetTemplateFile(string Id, string filename)
        {
            return _sf.GetAttachmentFile(Id, templatefilename, filename);
        }

        public DataReturn GetClauses()
        {
            return _sf.RunSOQL("select Id,Name,Description__c,Text__c,Concept__r.Id,Concept__r.Name,Concept__r.PlayBookInfo__c,Concept__r.PlayBookClient__c from " + this.ribbonclause + " order by Name");
        }

        public DataReturn GetClause(string Id)
        {
            return _sf.RunSOQL("select Id,Name,Description__c,Text__c,Concept__r.Id,Concept__r.Name,Concept__r.PlayBookInfo__c,Concept__r.PlayBookClient__c from " + this.ribbonclause + " where Id = '" + Id + "'");
        }


        public DataReturn GetClauseFile(string Id, string filename)
        {
            return _sf.GetAttachmentFile(Id, this.templatefilename, filename);
        }



        public DataReturn SaveClause(DataRow dr)
        {
            return _sf.Save(this.ribbonclause, dr);
        }

        public DataReturn DeleteClause(String Id)
        {
            return _sf.Delete(this.ribbonclause, Id);
        }

        public DataReturn CheckClause(string name)
        {
            name = Utility.FixUpSOQLString(name.Trim());
            return _sf.RunSOQL("SELECT Id from " + this.ribbonclause + " where Name = '" + name + "'");
        }


        public DataReturn SaveClauseFile(string Id, string Text, string filename)
        {
            DataReturn rtn;
            //Clause have to save the text as well
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Text__c", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Text__c"] = Utility.Truncate(Text, 131072);
            rtn = _sf.Save(this.ribbonclause, rw);


            if (rtn.success)
            {

                // get the last modified date
                DataReturn mdate = _sf.RunSOQL("select LastModifiedDate from " + this.ribbonclause + " where Id = '" + rtn.id + "'");

                //remember the clause id - more intrested in that than the attachment id
                string clauseid = rtn.id;
                rtn = _sf.SaveAttachmentFile(clauseid, templatefilename, filename);

                rtn.strRtn = mdate.dt.Rows[0][0].ToString();
                rtn.id = clauseid;
            }
            return rtn;
        }


        public DataReturn SaveClauseFromTemplateClause(DataRow dr)
        {
            //Save changes made to the tree view which is displaying the TemplateClause rather than the Clause
            //e.g. select Id,Name,Clause__r.Id,Clause__r.Name,Clause__r.Concept__r.Id,Clause__r.Concept__r.Name from RibbonTemplateClause__c
            //but want to save the changes to just the clause fields so create a new datatable
            //and save

            //Create a new datatable
            DataTable dt = new DataTable();
            foreach (DataColumn c in dr.Table.Columns)
            {
                if (c.ColumnName.StartsWith("Clause__r_"))
                {
                    string tname = c.ColumnName.Substring("Clause__r_".Length);
                    dt.Columns.Add(new DataColumn(tname, c.DataType));
                }
            }
            DataRow rw = dt.NewRow();
            foreach (DataColumn c in dt.Columns)
            {
                rw[c] = dr["Clause__r_" + c.ColumnName];
            }

            DataReturn rtn;
            rtn = _sf.Save(this.ribbonclause, rw);

            return rtn;
        }

        public DataReturn SaveTemplateClause(string Id, string Name, string TemplateId, string ClauseId, string Order, string DefaultSelection)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            if (Id == "") dt.Columns.Add(new DataColumn("Template__c", typeof(String)));
            if (Id == "") dt.Columns.Add(new DataColumn("Clause__c", typeof(String)));
            dt.Columns.Add(new DataColumn("Order__c", typeof(String)));
            dt.Columns.Add(new DataColumn("DefaultSelection__c", typeof(String)));

            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Name"] = Utility.Truncate(Name, 80);
            if (Id == "") rw["Template__c"] = TemplateId;
            if (Id == "") rw["Clause__c"] = ClauseId;
            rw["DefaultSelection__c"] = DefaultSelection;

            string doubleorder = "0";
            try
            {
                doubleorder = Convert.ToDouble(Order).ToString();
            }
            catch (Exception)
            {
                doubleorder = "0";
            }
            rw["Order__c"] = doubleorder;

            return _sf.Save(this.ribbontemplateclause, rw);
        }



        public DataReturn GetTemplateClauseCount(string TemplateId, string ConceptId)
        {
            return _sf.RunSOQL("SELECT Id FROM " + this.ribbontemplateclause + " where Clause__r.Concept__c = '" + ConceptId + "' and Template__c = '" + TemplateId + "'");
        }

        public DataReturn GetTemplateClauses(string TemplateId, string ConceptId)
        {

            string sql = "select Id,Name,Order__c,DefaultSelection__c,Template__c,Clause__r.Id,Clause__r.Name,Clause__r.Concept__r.Id,Clause__r.Concept__r.Name,Clause__r.Concept__r.Description__c,Clause__r.Concept__r.PlayBookInfo__c,Clause__r.Concept__r.PlayBookClient__c,Clause__r.Concept__r.AllowNone__c,Clause__r.LastModifiedDate from " + this.ribbontemplateclause
            + " where Template__c = '" + TemplateId + "' and Clause__r.Concept__r.Id = '" + ConceptId + "' order by Order__c";

            if (ConceptId == "")
            {
                sql = "select Id,Name,Order__c,DefaultSelection__c,Template__c,Clause__r.Id,Clause__r.Name,Clause__r.Description__c,Clause__r.Concept__r.Id,Clause__r.Concept__r.Name,Clause__r.Concept__r.Description__c,Clause__r.Concept__r.PlayBookInfo__c,Clause__r.Concept__r.PlayBookClient__c,Clause__r.Concept__r.AllowNone__c,Clause__r.LastModifiedDate from " + this.ribbontemplateclause
              + " where Template__c = '" + TemplateId + "' order by Order__c";
            }

            return _sf.RunSOQL(sql);

        }

        public DataReturn GetTemplateClause(string TemplateClauseId)
        {
            string sql = "select Id,Name,Order__c,DefaultSelection__c,Template__c,Clause__r.Id,Clause__r.Name,Clause__r.Concept__r.Id,Clause__r.Concept__r.Name,Clause__r.Concept__r.Description__c,Clause__r.Concept__r.PlayBookInfo__c,Clause__r.Concept__r.PlayBookClient__c,Clause__r.Concept__r.AllowNone__c,Clause__r.LastModifiedDate from " + this.ribbontemplateclause
            + " where Id = '" + TemplateClauseId + "'";

            return _sf.RunSOQL(sql);
        }

        public DataReturn GetTemplateClause(string TemplateId, string ClauseId)
        {
            string sql = "select Id from " + this.ribbontemplateclause
            + " where Template__c = '" + TemplateId + "'"
            + " and Clause__c = '" + ClauseId + "'";

            return _sf.RunSOQL(sql);
        }

        public DataReturn DeleteTemplateClause(string Id)
        {
            return _sf.Delete(this.ribbontemplateclause, Id);
        }


        public DataReturn GetConcepts()
        {
            return _sf.RunSOQL("select Id,Name,AllowNone__c,Description__c from " + this.ribbonconcept + " order by Name");
        }

        public DataReturn CheckConcept(string name)
        {
            name = Utility.FixUpSOQLString(name.Trim());
            return _sf.RunSOQL("SELECT Id from " + this.ribbonconcept + " where Name = '" + name + "'");
        }

        public DataReturn GetElements()
        {
            return _sf.RunSOQL("SELECT Id,Name,Label__c,Description__c,Type__c,Format__c,Options__c,DefaultValue__c FROM " + this.ribbonelement + " order by Name");
        }


        public DataReturn GetElements(string ClauseId)
        {
            return _sf.RunSOQL("SELECT Id,Name,Order__c,Element__r.Id,Element__r.Name,Element__r.Label__c,Element__r.Type__c,Element__r.Description__c,Element__r.Format__c,Element__r.Options__c,Element__r.DefaultValue__c,Clause__r.Id,Clause__r.Name,Clause__r.Concept__r.Name FROM " + ribbonclauseelement + " where Clause__c = '" + ClauseId + "' order by Order__c");
        }

        public DataReturn GetMultipleClauseElements(string clausefilter)
        {
            // if there are no clauses then just get the sql with the 1=2 - need to get the table structure
            string select = "SELECT Id,Name,Order__c,Element__r.Id,Element__r.Name,Element__r.Label__c,Element__r.Type__c,Element__r.Description__c,Element__r.Format__c,Element__r.Options__c,Element__r.DefaultValue__c,Clause__r.Id,Clause__r.Name,Clause__r.Concept__r.Name FROM " + ribbonclauseelement;
            if (clausefilter == "")
            {
                return _sf.RunSOQL(select + " where Clause__c='123456789012345678'");
            }
            else
            {
                return _sf.RunSOQL(select + " where Clause__c in " + clausefilter + " order by Order__c");
            }
        }


        public DataReturn GetElement(string ClauseId, String ElementId)
        {

            return _sf.RunSOQL("SELECT Id FROM " + this.ribbonclauseelement + "  where Clause__c = '" + ClauseId + "' and Element__c = '" + ElementId + "'");
        }


        public DataReturn GetElement(string Id)
        {
            return _sf.RunSOQL("SELECT Id,Name,Label__c,Description__c,Type__c,Format__c,Options__c,DefaultValue__c FROM " + this.ribbonelement + " where Id='" + Id + "'");
        }

        public DataReturn SaveElement(DataRow dr)
        {
            return _sf.Save(this.ribbonelement, dr);
        }

        public DataReturn SaveElementFromClauseElement(DataRow dr)
        {
            //Save the element values when loaded from the ClauseElement 
            //Create a new datatable
            DataTable dt = new DataTable();
            foreach (DataColumn c in dr.Table.Columns)
            {
                if (c.ColumnName.StartsWith("Element__r_"))
                {
                    string tname = c.ColumnName.Substring("Element__r_".Length);
                    dt.Columns.Add(new DataColumn(tname, c.DataType));
                }
            }
            DataRow rw = dt.NewRow();
            foreach (DataColumn c in dt.Columns)
            {
                rw[c] = dr["Element__r_" + c.ColumnName];
            }

            return _sf.Save(this.ribbonelement, rw);

        }


        public DataReturn DeleteElement(String Id)
        {
            return _sf.Delete(this.ribbonelement, Id);
        }

        public DataReturn DeleteClauseElement(string Id)
        {
            return _sf.Delete(this.ribbonclauseelement, Id);
        }

        public DataReturn SaveClauseElement(string Id, string Name, string ClauseId, string ElementId, string Order)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            dt.Columns.Add(new DataColumn("Order__c", typeof(String)));
            if (Id == "") dt.Columns.Add(new DataColumn("Clause__c", typeof(String)));
            if (Id == "") dt.Columns.Add(new DataColumn("Element__c", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Name"] = Name;
            rw["Order__c"] = Order;
            if (Id == "") rw["Clause__c"] = ClauseId;
            if (Id == "") rw["Element__c"] = ElementId;

            return _sf.Save(this.ribbonclauseelement, rw);
        }

        public DataReturn UpdateClauseElementOrder(string Id, string Order)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Order__c", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Order__c"] = Order;

            return _sf.Save(this.ribbonclauseelement, rw);
        }

        public DataReturn CheckElement(string name)
        {
            name = Utility.FixUpSOQLString(name.Trim());
            return _sf.RunSOQL("SELECT Id FROM " + this.ribbonelement + " where Name='" + name + "'");
        }

        public DataReturn UpdateTemplateClause(string Id, string Order, string DefaultSelection)
        {

            // check that the order is a number
            Decimal decord = 0;
            Decimal.TryParse(Order, out decord);

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Order__c", typeof(String)));
            dt.Columns.Add(new DataColumn("DefaultSelection__c", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Order__c"] = decord.ToString();
            rw["DefaultSelection__c"] = DefaultSelection;

            return _sf.Save(this.ribbontemplateclause, rw);
        }


        public DataReturn UpdateConceptAllowNone(string Id, bool? val)
        {
            if (val == null) val = false;

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("AllowNone__c", typeof(bool)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["AllowNone__c"] = val;

            return _sf.Save(this.ribbonconcept, rw);
        }


        // for the Clean Up - update the name of the Concept
        public DataReturn UpdateConceptName(string Id, string Name)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Name"] = Name;

            return _sf.Save(this.ribbonconcept, rw);
        }

        public DataReturn UpdateClauseName(string Id, string Name)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = Id;
            rw["Name"] = Name;

            return _sf.Save(this.ribbonclause, rw);

        }

        //Contract Objects ----------------------------------------------------------------------


        public DataReturn SaveVersion(string VersionId, string MatterId, string TemplateId, string Name, string Number)
        {
            // Save the Version - will create a new version if the VersionId is blank

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            dt.Columns.Add(new DataColumn("Version_Number__c", typeof(Double)));
            dt.Columns.Add(new DataColumn("Template__c", typeof(String)));

            string parenttable = "Matter__c";
            string table = this.version;

            // get it to work when its straight from a request 
            if (this.demoinstance == "general")
            {
                parenttable = "Request2__c";
            }
            else if (this.demoinstance == "isda")
            {
                parenttable = "Version__c";
                table = "Document__c";
            }

            if (VersionId == "") dt.Columns.Add(new DataColumn(parenttable, typeof(String)));

            DataRow rw = dt.NewRow();
            rw["Id"] = VersionId;
            rw["Name"] = Name;
            rw["Version_Number__c"] = (Number == "" ? "1" : Number);
            rw["Template__c"] = TemplateId;
            if (VersionId == "") rw[parenttable] = MatterId;

            return _sf.Save(table, rw);
        }
        //New PES
        public DataReturn CreateVersion(string VersionId, string MatterId, string TemplateId, string Name, string Number,DataRow dr)
        {
            // Save the Version - will create a new version if the VersionId is blank

                    DataTable dt = new DataTable();
                    dt.Columns.Add(new DataColumn("Id", typeof(String)));
                    dt.Columns.Add(new DataColumn("Name", typeof(String)));
                    dt.Columns.Add(new DataColumn("Version_Number__c", typeof(Double)));
                    dt.Columns.Add(new DataColumn("Template__c", typeof(String)));

                    string parenttable = "Matter__c";
                    string table = this.version;

                    // get it to work when its straight from a request 
                    if (this.demoinstance == "general")
                    {
                        parenttable = "Request2__c";
                    }
                    else if (this.demoinstance == "isda")
                    {
                        parenttable = "Version__c";
                        table = "Document__c";
                    }

                    if (VersionId == "") dt.Columns.Add(new DataColumn(parenttable, typeof(String)));
                  //  DataRow rw = dr;
                    DataRow rw = dt.NewRow();
                    rw["Id"] = VersionId;
                    rw["Name"] = Name;
                    rw["Version_Number__c"] = (Number == "" ? "1" : Number);
                    rw["Template__c"] = TemplateId;
                    if (VersionId == "") rw[parenttable] = MatterId;

                    return _sf.Save(table, rw);

            /* *
            for (int i = dr.Table.Columns.Count - 1; i >= 0; i--)
            {
                if (dr[i] == "" || dr[i] == null)
                {
                    dr.Table.Columns.RemoveAt(i);
                }
            }
            dr["Id"] = "";
            dr["Name"] = Name;
            dr["Version_Number__c"] = (Number == "" ? "1" : Number);
            dr["Template__c"] = TemplateId;
            if (VersionId == "") dr["matter__c"] = MatterId;
            return _sf.Save("Version__c", dr);
             */
        }
        // End PES

        public DataReturn SaveDocumentClause(string DocumentClauseId, string DocumentId, string ConceptId, string ClauseId, string Name, int Seq, string Text, bool unlock)
        {
            // TODO Truncs should be done by the Salesforce save automatically using the SF Definitions
            // just doing this till I switch the save over to the correct way

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            if (DocumentClauseId == "") dt.Columns.Add(new DataColumn(this.version, typeof(String)));
            dt.Columns.Add(new DataColumn("Concept__c", typeof(String)));
            dt.Columns.Add(new DataColumn("SelectedClause__c", typeof(String)));
            dt.Columns.Add(new DataColumn("Sequence__c", typeof(int)));
            dt.Columns.Add(new DataColumn("Text__c", typeof(String)));
            dt.Columns.Add(new DataColumn("StandardClause__c", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = DocumentClauseId;
            rw["Name"] = Utility.Truncate(Name, 80);
            if (DocumentClauseId == "") rw[this.version] = DocumentId;
            rw["Concept__c"] = ConceptId;
            rw["SelectedClause__c"] = ClauseId;
            rw["Sequence__c"] = Seq;
            rw["Text__c"] = Utility.Truncate(Text, 131072);
            rw["StandardClause__c"] = unlock ? "No" : "Yes";

            return _sf.Save(this.clause, rw);
        }


        public DataReturn GetDocumentFile(string Id, string filename)
        {
            return _sf.GetAttachmentFile(Id, contractfilename, filename);
        }

        public DataReturn SaveDocumentFile(string Id, string filename)
        {
            return _sf.SaveAttachmentFile(Id, this.contractfilename, filename);
        }

        public DataReturn GetDocumentClause(string documentid, string conceptid)
        {
            return _sf.RunSOQL("SELECT Id,SelectedClause__c,StandardClause__c FROM " + this.clause + " where Concept__c ='" + conceptid + "' and " + this.version + " = '" + documentid + "'");
        }


        //Get them all
        public DataReturn GetDocumentClause(string documentid)
        {
            return _sf.RunSOQL("SELECT Id,Concept__c,SelectedClause__c,StandardClause__c,Name,Sequence__c,Text__c,Version__c FROM " + this.clause + " where " + this.version + " = '" + documentid + "' order by Concept__c");
        }



        public DataReturn GetVersions()
        {
            if (this.demoinstance == "isda")
            {
                return _sf.RunSOQL("SELECT Id,Name,Version_Number__c,Template__c,Template__r.Name,Template__r.PlaybookLink__c,Version__c,Version__r.Name FROM " + this.version + " where Template__c <> '' order by Name");
            }
            else if (this.demoinstance == "general")
            {
                return _sf.RunSOQL("SELECT Id,Name,Version_Number__c,Template__c,Template__r.Name,Template__r.PlaybookLink__c,Request2__c,Request2__r.Name  FROM " + this.version + " where Template__c <> '' order by Name");
            }
            else
            {
                return _sf.RunSOQL("SELECT Id,Name,Version_Number__c,Template__c,Template__r.Name,Template__r.PlaybookLink__c,Matter__c,Matter__r.Name FROM " + this.version + " where Template__c <> '' order by Name");
            }
        }

        public DataReturn GetVersion(string Id)
        {

            if (this.demoinstance == "isda")
            {
                return _sf.RunSOQL("SELECT Id,Name,Version_Number__c,Template__c,Template__r.Name,Template__r.PlaybookLink__c,Version__c,Version__r.Name FROM " + this.version + " where Id = '" + Id + "'");
            }
            else if (this.demoinstance == "general")
            {
                return _sf.RunSOQL("SELECT Id,Name,Version_Number__c,Template__c,Template__r.Name,Template__r.PlaybookLink__c,Request2__c,Request2__r.Name FROM " + this.version + " where Id = '" + Id + "'");
            }
            else
            {
                return _sf.RunSOQL("SELECT Id,Name,Version_Number__c,Template__c,Template__r.Name,Template__r.PlaybookLink__c,Matter__c,Matter__r.Name FROM  " + this.version + " where Id = '" + Id + "'");
            }
        }


        public DataReturn SaveContract(DataRow dr)
        {
            return _sf.Save(this.version, dr);
        }


        public DataReturn SaveDocumentClauseElement(string DocumentClauseElementId, string Name, string DocumentClauseId, string DocumentId, string TemplateElementId, string Value, string FormattedText)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            if (DocumentClauseElementId == "") dt.Columns.Add(new DataColumn(this.version, typeof(String)));
            if (DocumentClauseElementId == "") dt.Columns.Add(new DataColumn(this.clause, typeof(String)));
            dt.Columns.Add(new DataColumn("RibbonElement__c", typeof(String)));
            dt.Columns.Add(new DataColumn("Value__c", typeof(String)));
            dt.Columns.Add(new DataColumn("FormattedText__c", typeof(String)));

            DataRow rw = dt.NewRow();
            rw["Id"] = DocumentClauseElementId;
            rw["Name"] = Utility.Truncate(Name, 80);
            if (DocumentClauseElementId == "") rw[this.version] = DocumentId;
            if (DocumentClauseElementId == "") rw[this.clause] = DocumentClauseId;
            rw["RibbonElement__c"] = TemplateElementId;
            rw["Value__c"] = Value;
            rw["FormattedText__c"] = FormattedText;

            return _sf.Save(this.element, rw);
        }

        public DataReturn SaveDocumentClauseElement(DataRow dr)
        {
            return _sf.Save(this.element, dr);
        }

        public DataReturn GetDocumentElements(string DocumentId)
        {
            return _sf.RunSOQL("SELECT Id,Name,Value__c,RibbonElement__c FROM Element__c where " + this.version + " = '" + DocumentId + "' order by Name");
        }

        public DataReturn GetDocumentClauseElements(string ClauseId)
        {
            return _sf.RunSOQL("SELECT Id,Name,Value__c,RibbonElement__c,FormattedText__c,Version__c,Clause__c FROM Element__c where Clause__c = '" + ClauseId + "'");
        }


        public DataReturn GetConcept(string ConceptId)
        {
            string sql = "select Id,Name,PlayBookInfo__c,PlayBookClient__c from " + this.ribbonconcept
            + " where Id = '" + ConceptId + "'";
            return _sf.RunSOQL(sql);
        }

        public DataReturn SaveConcept(string ConceptId, string Info, string Client)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            if (Info != null) dt.Columns.Add(new DataColumn("PlayBookInfo__c", typeof(String)));
            if (Client != null) dt.Columns.Add(new DataColumn("PlayBookClient__c", typeof(String)));
            DataRow rw = dt.NewRow();
            rw["Id"] = ConceptId;

            if (Info != null) rw["PlayBookInfo__c"] = Utility.Truncate(Info, 131072);
            if (Client != null) rw["PlayBookClient__c"] = Utility.Truncate(Client, 131072);

            return _sf.Save(this.ribbonconcept, rw);
        }

        public DataReturn SaveConcept(string ConceptId, string Name, string Description, string Info, string Client, bool? AllowNone)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(String)));
            dt.Columns.Add(new DataColumn("Name", typeof(String)));
            dt.Columns.Add(new DataColumn("Description__c", typeof(String)));

            if (Info != null) dt.Columns.Add(new DataColumn("PlayBookInfo__c", typeof(String)));
            if (Client != null) dt.Columns.Add(new DataColumn("PlayBookClient__c", typeof(String)));
            dt.Columns.Add(new DataColumn("AllowNone__c", typeof(bool)));
            DataRow rw = dt.NewRow();
            rw["Id"] = ConceptId;
            rw["Name"] = Utility.Truncate(Name, 80);
            rw["Description__c"] = Utility.Truncate(Description, 32768);
            if (Info != null) rw["PlayBookInfo__c"] = Utility.Truncate(Info, 131072);
            if (Client != null) rw["PlayBookClient__c"] = Utility.Truncate(Client, 131072);

            if (AllowNone == null) AllowNone = false;
            rw["AllowNone__c"] = AllowNone;

            return _sf.Save(this.ribbonconcept, rw);

        }

        public DataReturn SaveConcept(DataRow rw)
        {
            return _sf.Save(this.ribbonconcept, rw);
        }

        public DataReturn DeleteConcept(String Id)
        {
            return _sf.Delete(this.ribbonconcept, Id);
        }


        /* Generic SalesForce Edit routines */

        public bool HasLibraryObjects()
        {
            //Check if this instance has the Clause Librarty Objects - just check for RibbonTemplates and if its there assume the others are

            if (_sf._allSObjects.ContainsKey(this.ribbontemplate))
            {
                return true;
            }

            return false;
        }

        public DataReturn LoadDefinitions()
        {
            string[] sObj = Globals.ThisAddIn.GetSettings().GetSetting("", "SObjects").Split('|');
            return _sf.LoadDefinitions(sObj);

        }

        public sfPartner.DescribeLayoutResult GetLayout(string name)
        {
            return _sf.GetDefinitionLayout(name);
        }

        public sfPartner.DescribeSObjectResult GetSObject(string name)
        {
            return _sf.GetDefinitionSObject(name);
        }

        public sfPartner.DescribeSearchLayoutResult GetSearch(string name)
        {
            return _sf.GetDefinitionSearch(name);
        }


        public sfPartner.Field GetField(string sobject, string field)
        {
            sfPartner.DescribeSObjectResult dsr = GetSObject(sobject);
            for (int j = 0; j < dsr.fields.Length; j++)
            {
                if (dsr.fields[j].name == field) return dsr.fields[j];
            }

            return null;
        }


        public DataReturn GetData(SForceEdit.SObjectDef sobj)
        {
            return _sf.RunSOQL(sobj);
        }

        public DataReturn Save(SForceEdit.SObjectDef sobj, DataRow r)
        {
            return _sf.Save(sobj, r);
        }

        public sfPartner.DescribeGlobalSObjectResult GetSObjectDef(string name)
        {
            return _sf.GetSObjectDef(name);
        }

        public string GetUrlForNonLoaded(string name)
        {
            return _sf.GetUrlForNonLoaded(name);
        }


        public string GetSessionId()
        {
            return _sf.GetSessionId();
        }

        public string GetURL()
        {
            return _sf.GetURL();

        }

        public string GetPartnerURL()
        {
            return _sf.GetPartnerURL();
        }

        public DataReturn AttachFile(string ParentId, string AttachmentName, string File)
        {
            DataReturn dr = _sf.SaveAttachmentFile(ParentId, AttachmentName, File);
            return dr;
        }
        public DataReturn UpdateFile(string Id, string AttachmentName, string File)
        {
            DataReturn dr = _sf.UpdateAttachmentFile(Id, AttachmentName, File);
            return dr;
        }

        public DataReturn OpenFile(string Id, string AttachName)
        {
            //temp file
            string file = GetTempFilePath(AttachName);
            DataReturn dr = _sf.GetAttachmentFile(Id, file);
            dr.strRtn = file;
            dr.id = Id;
            return dr;
        }

        public DataReturn DeleteFile(string Id)
        {
            DataReturn dr = _sf.Delete("Attachment", Id);
            return dr;
        }

        public DataReturn GetStaticResource(string Name)
        {
            DataReturn dr = _sf.GetStaticResource(Name);
            return dr;
        }


        public DataReturn Exec(string Action, string Obj, string Id)
        {
            DataReturn dr = _sf.ExecRibbonCall(Action, Obj, Id);
            return dr;
        }

        // Code PES - changed method as public
        public string GetTempFilePath(string AttachName)
        {
            //Generate a temp file and save the doc there
            string temppath = System.IO.Path.GetTempPath();
            string filename = "";
            int fcount = 0;
            while (filename == "")
            {
                if (System.IO.File.Exists(temppath + AttachName))
                {
                    try
                    {
                        // Russel McNeill - 13 Oct - oops was checking just AttachName
                        // so was getting the default path - so was not deleting the temp file
                        // but could delete the original
                        System.IO.File.Delete(temppath + AttachName);
                        filename = AttachName;
                    }
                    catch (Exception)
                    {
                        fcount++;
                        AttachName = System.IO.Path.GetFileNameWithoutExtension(AttachName) +
                                   "_" + fcount.ToString() + System.IO.Path.GetExtension(AttachName);
                    }
                }
                else
                {
                    filename = AttachName;
                }
            }
            return temppath + filename;
        }



        public DataReturn GetVersionMax(string MatterId)
        {
            // not totally sure about this - this give us the max count 
            // so as long as the version numbers are kept up to date should work
            DataReturn dr = _sf.RunSOQL("select max(Version_Number__c) from Version__c where Matter__c = '" + MatterId + "' and IsDeleted = false");
            return dr;
        }

        // data routines for the Compare function 
        public DataReturn GetVersionFromMatter(string MatterId)
        {
            DataReturn dr = _sf.RunSOQL("select Id,Name,Version_Number__c from Version__c where Matter__c = '" + MatterId + "' and IsDeleted = false order by Version_Number__c desc");
            return dr;
        }

        public DataReturn GetVersionAttachments(string VersionId)
        {
            // not totally sure about this - this give us the max count 
            // so as long as the version numbers are kept up to date should work
            DataReturn dr = _sf.RunSOQL("SELECT Id,Name FROM Attachment where ParentId='" + VersionId + "'");
            return dr;
        }

        //New PES
        public DataReturn GetVersionAllAttachments(string VersionId)
        {
            DataReturn dr = _sf.RunSOQL("SELECT Id,Name,Body FROM Attachment where ParentId='" + VersionId + "'");
            return dr;
        }

        public void saveAttachmentstoSF(string ParentId, string AttachmentName, string Xml)
        {
            //code
            _sf.CloneAttachmentFile(ParentId, AttachmentName, Xml);
        }
        //END PES
        public string GetInstanceInfo()
        {
            if (_sf._loggedin)
            {
                return _sf.GetInstanceInfo();
            }
            else
            {
                return "Not Logged In";
            }
        }

        public string GetUserInfo()
        {
            if (_sf._loggedin)
            {
                return _sf.GetUserInfo();
            }
            else
            {
                return "";
            }
        }


        // Had a one off Cipher issue have to reload the playbook Info from UAT
        public DataReturn GetConceptsPlaybookInfo()
        {
            return _sf.RunSOQL("select Id,Name,PlayBookInfo__c,PlayBookClient__c from " + this.ribbonconcept + " order by Name");
        }


    }
}

