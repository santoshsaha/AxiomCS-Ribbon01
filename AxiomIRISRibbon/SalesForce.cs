using AxiomIRISRibbon.sfPartner;
using AxiomIRISRibbon.sfRibbon;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Services.Protocols;
using System.Data;
using System.Collections;
using System.Text.RegularExpressions;

namespace AxiomIRISRibbon
{
    //Class to handle the interactions with SalesForce
    //Could make this more generic and use a dispatcher to pick
    //the access class to allow us to have difrent back ends - might come back to that

    public class SalesForce
    {
        private SforceService _binding;
        private Axiom_RibbonControllerService _ribbonBinding;

        private string _metaurl;
        public bool _loggedin;

        private string _userid;
        private string _url;
        public string _instancedesc;

        public Dictionary<string, DescribeGlobalSObjectResult> _allSObjects;
        public Dictionary<string, sfPartner.DescribeSObjectResult> _describeSObject;
        public Dictionary<string, sfPartner.DescribeLayoutResult> _describeLayout;
        public Dictionary<string, sfPartner.DescribeSearchLayoutResult> _describeSearch;

        
        public SalesForce()
        {
            _loggedin = false;
            _instancedesc = "";
        }

        

        public string Login(string Username,string Password,string Token,string Url,string InstanceDesc)
        {
            _instancedesc = InstanceDesc;

            // Create a service object 
            _binding = new SforceService();

            // Timeout after a minute 
            _binding.Timeout = 60000;

            // to get Fiddler (http://www.telerik.com/fiddler) to work - this can be helpful when trying to dubug Soap API issues
            // 
            // System.Net.WebProxy wp = new System.Net.WebProxy("http://127.0.0.1:8888", false);
            // _binding.Proxy = wp;

            // don't relly on the default binding url cause we want to set the API version
            if (Url == "") Url = "https://login.salesforce.com";

            try
            {
                if (Url != "")
                {
                    //if its just the web site address and not the soap endpoint add that in
                    if (!Url.Contains("services/Soap/"))
                    {
                        // don't use the default API version - use the defined one                       
                        LocalSettings l = Globals.ThisAddIn.GetLocalSettings();
                        Url += "/services/Soap/u/" + l.SoapVersion + "/";
                    }
                    _binding.Url = Url;
                }
            } catch (Exception e)
            {
                return "Problem with the URL: " + e.Message;
            }

            // Try logging in   
            LoginResult lr;
            try
            {
                lr = _binding.login(Username.Trim(), Password.Trim() + Token.Trim());
            }
            catch (SoapException e)
            {
                return e.Code + " " + e.Message;
            }
            catch (Exception e)
            {
                return e.Message;
            }


            // Check if the password has expired 
            if (lr.passwordExpired)
            {
                return "Password has expired - please login to Salesforce and update";
            }

            String authEndPoint = _binding.Url;
            _binding.Url = lr.serverUrl;

            Uri sforceurl = new Uri(lr.serverUrl);
            _url = sforceurl.Scheme + "://" + sforceurl.Host;

            //remember the meta end point may need later
            _metaurl = lr.metadataServerUrl;
            
            //remember who we are
            _binding.SessionHeaderValue = new sfPartner.SessionHeader();
            _binding.SessionHeaderValue.sessionId = lr.sessionId;
 
            _loggedin = true;
            _userid = lr.userId;


            // Get the url for the Ribbon web service binding
            _ribbonBinding = new Axiom_RibbonControllerService();
            _ribbonBinding.SessionHeaderValue = new sfRibbon.SessionHeader();
            _ribbonBinding.SessionHeaderValue.sessionId = lr.sessionId;

            int idx1 = lr.serverUrl.IndexOf(@"/services/");
            int idx2 = _ribbonBinding.Url.IndexOf(@"/services/");
            if (idx1 == 0 || idx2 == 0)
            {
                return "Pproblem with the urls";
            }
            _ribbonBinding.Url = lr.serverUrl.Substring(0, idx1) + _ribbonBinding.Url.Substring(idx2);
            

            return "";
        
        }



        public string Login(string token, string partnerUrl, string metaUrl, string InstanceDesc)
        {

            _instancedesc = InstanceDesc;

            // Create a service object 
            _binding = new SforceService();

            // Timeout after a minute 
            _binding.Timeout = 60000;

            // get the version from the default url
            int i1 = this._binding.Url.LastIndexOf('/') + 1;
            string version = this._binding.Url.Substring(i1);

            this._binding.Url = partnerUrl.Replace("/c/", "/u/");
            this._metaurl = metaUrl;

            // remember who we are
            this._binding.SessionHeaderValue = new sfPartner.SessionHeader();
            this._binding.SessionHeaderValue.sessionId = token;
            this._loggedin = true;


            // Setup and Get the url for the Ribbon web service binding
            _ribbonBinding = new Axiom_RibbonControllerService();
            _ribbonBinding.SessionHeaderValue = new sfRibbon.SessionHeader();
            _ribbonBinding.SessionHeaderValue.sessionId = token;
            int idx1 = this._binding.Url.IndexOf(@"/services/");
            int idx2 = _ribbonBinding.Url.IndexOf(@"/services/");
            if (idx1 == 0 || idx2 == 0)
            {
                return "problem with the urls";
            }
            _ribbonBinding.Url = this._binding.Url.Substring(0, idx1) + _ribbonBinding.Url.Substring(idx2);


            // check we are logged in
            string ok = "";
            try
            {
                this._userid = _binding.getUserInfo().userId;
                this._loggedin = true;
                ok = "";
            }
            catch (Exception e)
            {
                this._loggedin = false;
                ok = e.Message;
            }

            return ok;
        }

        public void Logout()
        {
            try
            {
                _binding.logout();
            }
            catch (Exception)
            {

            }
            return;
        }


        public string GetSessionId(){
            return _binding.SessionHeaderValue.sessionId;
        }

        public string GetURL()
        {
            return _url;
        }

        public string GetPartnerURL()
        {
            return _binding.Url;
        }

        public DataReturn GetPickListValues(string sObject,string fName){

            DataReturn dr = new DataReturn();
            DataTable dt = dr.dt;

            
            System.Data.DataColumn c = new DataColumn("Value", typeof(String));
            dt.Columns.Add(c);
            try
            {
                DescribeSObjectResult[] dsrArray = _binding.describeSObjects(new string[] { sObject });
                DescribeSObjectResult dsr = dsrArray[0];

                for (int i = 0; i < dsr.fields.Length; i++)
                {
                    Field field = dsr.fields[i];
                    if (field.name == fName)
                    {
                        if (field.type.Equals(fieldType.picklist))
                        {
                            for (int j = 0; j < field.picklistValues.Length; j++)
                            {
                                DataRow rw = dt.NewRow();
                                rw["Value"] = field.picklistValues[j].value;
                                dt.Rows.Add(rw);
                            }
                        }
                    }
                }

            } catch(Exception e){
                dr.success = false;
                dr.errormessage = e.Message;                
            }
            return dr;
        }

        public string GetUser()
        {
            return _binding.getUserInfo().userFullName;            
        }

        public string GetInstanceInfo()
        {
            string x = "";
            if (_instancedesc != "")
            {
                x = _instancedesc;
            }
            else
            {
                x = _binding.getUserInfo().organizationName;
            }
            return x;
        }

        public string GetUserInfo()
        {
            string x = "";
            x += _binding.getUserInfo().userFullName + " (" + this.GetUserProfile() + ")";
            return x;
        }

        public string GetUserId()
        {
            return _userid;
        }

        public string GetUserProfile()
        {
            string profileid = _binding.getUserInfo().profileId;
            DataReturn dr = this.RunSOQL("select Name from Profile where Id='" + profileid + "'");
            
            if (!dr.success) return "";

            if (dr.dt.Rows.Count == 1)
            {
                return dr.dt.Rows[0][0].ToString();
            }
            else
            {
                return "";
            }
            
        }


        //Add Columns to the DataSet - need to call recursivley for those pesky relationships
        private void AddColumn(DataTable dt, System.Xml.XmlElement x,string ParentName)
        {

            if (ParentName != "") ParentName += "_";

            if (x.HasAttributes && x.Attributes["xsi:type"] != null && x.Attributes["xsi:type"].Value == "sf:sObject")
            {
                string tempname = ParentName + x.LocalName;
                for (int k = 0; k < x.ChildNodes.Count; k++)
                {
                    System.Xml.XmlElement xchild = ((System.Xml.XmlElement)x.ChildNodes[k]);
                    if (xchild.HasAttributes && xchild.Attributes["xsi:type"] != null && xchild.Attributes["xsi:type"].Value == "sf:sObject")
                    {
                        AddColumn(dt, xchild, tempname);
                    }
                    else
                    {
                        if (!dt.Columns.Contains(tempname + "_" + xchild.LocalName))
                        {
                            System.Data.DataColumn c = new DataColumn(tempname + "_" + xchild.LocalName, typeof(String));
                            dt.Columns.Add(c);
                        }
                    }

                }
            }
            else
            {
                if (!dt.Columns.Contains(x.LocalName))
                {
                    System.Data.DataColumn c = new DataColumn(x.LocalName, typeof(String));
                    dt.Columns.Add(c);
                }
            }
        }

        //Add Data to the Dataset - need to call recursilvely
        private void AddData(DataRow rw, System.Xml.XmlElement x, string ParentName)
        {
            if (ParentName != "") ParentName += "_";

            if (x.HasAttributes && x.Attributes["xsi:type"] != null && x.Attributes["xsi:type"].Value == "sf:sObject")
            {
                string tempname = ParentName + x.LocalName;
                for (int k = 0; k < x.ChildNodes.Count; k++)
                {
                    System.Xml.XmlElement xchild = ((System.Xml.XmlElement)x.ChildNodes[k]);
                    if (xchild.HasAttributes && xchild.Attributes["xsi:type"] != null && xchild.Attributes["xsi:type"].Value == "sf:sObject")
                    {
                        AddData(rw, xchild, tempname);
                    }
                    else
                    {
                        if(rw.Table.Columns.Contains(tempname + "_" + xchild.LocalName)){
                            rw[tempname + "_" + xchild.LocalName] = xchild.InnerText;
                        }
                    }
                }
            }
            else
            {
                if (rw.Table.Columns.Contains(x.LocalName))
                {
                    rw[x.LocalName] = x.InnerText;
                }
            }

        }

        public DataReturn RunSOQL(string soqlQuery)
        {
            //todo - add page handling
            //for now get everything

            DataReturn dr = new DataReturn();
            DataTable dt = dr.dt;

            try
            {
                QueryResult qr = _binding.query(soqlQuery);
                bool done = false;
                bool first = true;

                if (qr.size > 0)
                {
                    while (!done)
                    {
                        sObject[] records = qr.records;
                        for (int i = 0; i < qr.records.Length; i++)
                        {

                            if (first)
                            {
                                //Build the datatable
                                for (int j = 0; j < records[i].Any.Length; j++)
                                {
                                    AddColumn(dt, records[i].Any[j],"");
                                }
                                first = false;
                            }

                            DataRow rw = dt.NewRow();
                            for (int j = 0; j < records[i].Any.Length; j++)
                            {
                                AddData(rw, records[i].Any[j],"");
                            }

                            dt.Rows.Add(rw);

                        }

                        if (qr.done)
                        {
                            done = true;
                        }
                        else
                        {
                            qr = _binding.queryMore(qr.queryLocator);
                        }
                    }
                }
                else
                {
                   //Still need to create the table so have the template when creating a new one
                   //just create string fields for each of the select fields - don't mind the amatuer hour parsing!

                    string temp = soqlQuery.Trim();
                    temp = temp.Substring("select".Length, temp.Length - "select".Length);
                    temp = temp.Substring(0, temp.IndexOf(" from ",StringComparison.CurrentCultureIgnoreCase));
                    temp = temp.Trim();
                    string[] fieldnames = temp.Split(',');

                    foreach (string f in fieldnames)
                    {
                        dt.Columns.Add(new DataColumn(f.Replace(".","_").Trim(), typeof(String)));
                    }

                }
            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }

            if (dt.Columns.Contains("Id")) dt.PrimaryKey = new DataColumn[] {dt.Columns["Id"]};

            return dr;
        }


        //Given a DataRow, update or Create the SalesForce Object
        //Assuming that we have just one row, easy to change to handle multiples
        public DataReturn Save(string sObjectName, DataRow dRow)
        {
            
            DataReturn dr = new DataReturn();

            sObject s = new sObject();
            s.type = sObjectName;
            string id = "";
            List<string> fieldsToNull = new List<string>();

            if (dRow["Id"] == null || dRow["Id"].ToString() == "")
            {
                //new
                int fldCount = dRow.Table.Columns.Count - 1;
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                System.Xml.XmlElement[] o = new System.Xml.XmlElement[fldCount];

                fldCount = 0;
                foreach (DataColumn dc in dRow.Table.Columns)
                {
                    if (dc.ColumnName == "Id")
                    {
                        //Nothing!
                    }
                    else if (dc.ColumnName == "type" || dc.ColumnName == "LastModifiedDate" || dc.ColumnName == "CreatedDate")
                    {
                        //don't do anything - this happens when we have the type field from a join
                    }
                    else
                    {
                        if (dc.ColumnName.Contains("__r_"))
                        {
                            if (dc.ColumnName.EndsWith("Id"))
                            {

                                string tn = dc.ColumnName;
                                tn = tn.Substring(0, tn.IndexOf("__r_"));
                                //Concept__r_Id becomes Concept__c
                                tn += "__c";

                                o[fldCount] = doc.CreateElement(tn);
                                o[fldCount].InnerText = CleanUpXML(dRow[dc.ColumnName].ToString());
                                fldCount++;
                            }
                            //Otherwise do nothing
                        }
                        else
                        {
                            o[fldCount] = doc.CreateElement(dc.ColumnName);
                            o[fldCount].InnerText = CleanUpXML(dRow[dc.ColumnName].ToString());
                            fldCount++;
                        }
                    }
                }

                try
                {

                    s.Any = Utility.SubArray<System.Xml.XmlElement>(o, 0, fldCount);
                    SaveResult[] sr = _binding.create(new sObject[] { s });


                    for (int j = 0; j < sr.Length; j++)
                    {
                        if (sr[j].success)
                        {
                            dr.id = sr[j].id;
                        }
                        else
                        {
                            dr.success = false;
                            for (int i = 0; i < sr[j].errors.Length; i++)
                            {
                                dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i].message;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    dr.success = false;
                    dr.errormessage = ex.Message;
                }
            }
            else
            {
                //update
                int fldCount = dRow.Table.Columns.Count;
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();

                System.Xml.XmlElement[] o = new System.Xml.XmlElement[fldCount];

                fldCount = 0;
                foreach (DataColumn dc in dRow.Table.Columns)
                {
                    if (dc.ColumnName == "Id")
                    {
                        s.Id = dRow[dc.ColumnName].ToString();
                    }
                    else if (dc.ColumnName == "type" || dc.ColumnName == "LastModifiedDate" || dc.ColumnName == "CreatedDate")
                    {
                        //don't do anything - this happens when we have the type field from a join
                    }
                    else
                    {
                        //For relations - ignore all the other fields except the _Id one
                        //e.g. "Concept__r_Name" "Concept__r_Id" - ignore all but Id

                        //TODO: won't work Nested! need to try this out with a realation of a relation

                        if (dc.ColumnName.Contains("__r_"))
                        {
                            if (dc.ColumnName.EndsWith("Id"))
                            {

                                string tn = dc.ColumnName;
                                tn = tn.Substring(0, tn.IndexOf("__r_"));
                                //Concept__r_Id becomes Concept__c
                                tn += "__c";

                                string val = CleanUpXML(dRow[dc.ColumnName].ToString());
                                if (val == "")
                                {
                                    fieldsToNull.Add(dc.ColumnName);
                                }
                                else
                                {
                                    o[fldCount] = doc.CreateElement(tn);
                                    o[fldCount].InnerText = val;
                                    fldCount++;
                                }

                            }
                            //Otherwise do nothing
                        }
                        else
                        {
                            string val = CleanUpXML(dRow[dc.ColumnName].ToString());
                            if(val==""){
                                fieldsToNull.Add(dc.ColumnName);
                            } else{    
                                o[fldCount] = doc.CreateElement(dc.ColumnName);
                                o[fldCount].InnerText = val;
                                fldCount++;
                            }
                        }

                    }
                }

                try
                {
                    s.fieldsToNull = fieldsToNull.ToArray();
                    s.Any = Utility.SubArray<System.Xml.XmlElement>(o, 0, fldCount);
                    SaveResult[] sr = _binding.update(new sObject[] { s });

                    for (int j = 0; j < sr.Length; j++)
                    {
                        Console.WriteLine("\nItem: " + j);
                        if (sr[j].success)
                        {
                            dr.id = sr[j].id;
                        }
                        else
                        {
                            dr.success = false;
                            for (int i = 0; i < sr[j].errors.Length; i++)
                            {
                                dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i].message;
                            }
                        }
                    }


                }
                catch (Exception ex)
                {
                    dr.success = false;
                    dr.errormessage = ex.Message;
                }

            }
            return dr;
        }



        public DataReturn Delete(string sObjectName, string id)
        {
            DataReturn dr = new DataReturn();

            DeleteResult[] drslt = _binding.delete(new string[]{id});
            Globals.Ribbons.Ribbon1.SFDebug("Delete", "Delete:" + id);

            for (int j = 0; j < drslt.Length; j++)
                {
                    DeleteResult deleteResult = drslt[j];
                    if (deleteResult.success)
                    {
                        dr.id = deleteResult.id;                        
                    }
                    else
                    {
                        Error[] errors = deleteResult.errors;
                        for (int k = 0; k < errors.Length; k++)
                        {
                            dr.errormessage += (dr.errormessage == "" ? "" : ",") + errors[k];
                        }
                    }
                }

                return dr;
        }
//New PES
        public DataReturn CloneAttachmentFile(string ParentId, string AttachmentName, string Xml)
        {
            DataReturn dr = new DataReturn();

            string id = "";

            String soqlQuery = "SELECT Id FROM Attachment where ParentId='" + ParentId + "' and Name='" + AttachmentName + "' order by LastModifiedDate desc limit 1";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        id = records[i].Any[0].InnerText;
                    }
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }


            sObject attach = new sObject();
            attach.type = "Attachment";
            System.Xml.XmlElement[] o;
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            SaveResult[] sr;

            if (id == "")
            {
                // Create the attacchments fields
                o = new System.Xml.XmlElement[4];
                doc = new System.Xml.XmlDocument();
                o[0] = doc.CreateElement("Name");
                o[0].InnerText = AttachmentName;

                o[1] = doc.CreateElement("isPrivate");
                o[1].InnerText = "false";

                o[2] = doc.CreateElement("ParentId");
                o[2].InnerText = ParentId;

                o[3] = doc.CreateElement("Body");
                byte[] data = Convert.FromBase64String(Xml);
                o[3].InnerText = Convert.ToBase64String(data);

                attach.Any = o;
                sr = _binding.create(new sObject[] { attach });

                for (int j = 0; j < sr.Length; j++)
                {
                    if (sr[j].success)
                    {
                        id = sr[j].id;
                    }
                    else
                    {
                        for (int i = 0; i < sr[j].errors.Length; i++)
                        {
                            dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];
                        }
                    }
                }
            }
            else
            {
                // Update the attacchments fields
                doc = new System.Xml.XmlDocument();
                o = new System.Xml.XmlElement[1];
                o[0] = doc.CreateElement("Body");
                //  o[0].InnerText = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(Xml));
                byte[] data = Convert.FromBase64String(Xml);
                o[0].InnerText = Convert.ToBase64String(data);


                attach.Any = o;
                attach.Id = id;
                sr = _binding.update(new sObject[] { attach });

                for (int j = 0; j < sr.Length; j++)
                {
                    if (sr[j].success)
                    {
                        id = sr[j].id;
                    }
                    else
                    {
                        for (int i = 0; i < sr[j].errors.Length; i++)
                        {
                            dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];
                        }
                    }
                }
            }

            dr.id = id;

            return dr;
        }
//End PES
        //Save attachment, using for the Word XML save for the templates and clauses
        //check if there is one already and either update or create
        public DataReturn SaveAttachment(string ParentId, string AttachmentName, string Xml)
        {
            DataReturn dr = new DataReturn();

            string id = "";

            String soqlQuery = "SELECT Id FROM Attachment where ParentId='" + ParentId + "' and Name='" + AttachmentName + "' order by LastModifiedDate desc limit 1";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        id = records[i].Any[0].InnerText;
                    }
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }


            sObject attach = new sObject();
            attach.type = "Attachment";
            System.Xml.XmlElement[] o;
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            SaveResult[] sr;

            if (id == "")
            {
                // Create the attacchments fields
                o = new System.Xml.XmlElement[4];
                doc = new System.Xml.XmlDocument();
                o[0] = doc.CreateElement("Name");
                o[0].InnerText = AttachmentName;

                o[1] = doc.CreateElement("isPrivate");
                o[1].InnerText = "false";

                o[2] = doc.CreateElement("ParentId");
                o[2].InnerText = ParentId;

              //  o[3] = doc.CreateElement("Body");
            //    o[3].InnerText = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(Xml));
                o[3] = doc.CreateElement("Body");
                byte[] data = Convert.FromBase64String(Xml);
                o[3].InnerText =Convert.ToBase64String( data);

                attach.Any = o;
                sr = _binding.create(new sObject[] { attach });

                for (int j = 0; j < sr.Length; j++)
                {
                    if (sr[j].success)
                    {
                        id = sr[j].id;
                    }
                    else
                    {
                        for (int i = 0; i < sr[j].errors.Length; i++)
                        {
                            dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];                            
                        }
                    }
                }
            }
            else
            {
                // Update the attacchments fields
                doc = new System.Xml.XmlDocument();
                o = new System.Xml.XmlElement[1];
                o[0] = doc.CreateElement("Body");
              //  o[0].InnerText = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(Xml));

                attach.Any = o;
                attach.Id = id;
                sr = _binding.update(new sObject[] { attach });

                for (int j = 0; j < sr.Length; j++)
                {
                    if (sr[j].success)
                    {
                        id = sr[j].id;
                    }
                    else
                    {
                        for (int i = 0; i < sr[j].errors.Length; i++)
                        {
                            dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];
                        }
                    }
                }
            }

            dr.id = id;

            return dr;
        }

        public DataReturn GetAttachment(string Id, string AttachmentName)
        {
            DataReturn dr = new DataReturn();

            String soqlQuery = "SELECT Body FROM Attachment where ParentId='" + Id + "' and Name='" + AttachmentName + "'  order by LastModifiedDate desc limit 1";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);
                Globals.Ribbons.Ribbon1.SFDebug("Get Attachment", soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        dr.strRtn = records[i].Any[0].InnerText;
                    }
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }

            if (dr.strRtn != "")            
            {
                dr.strRtn = System.Text.Encoding.Unicode.GetString(Convert.FromBase64String(dr.strRtn));
            }

            return dr;
        }


        public DataReturn GetAttachmentFile(string Id, string AttachmentName,string filename)
        {
            DataReturn dr = new DataReturn();
            byte[] b;
            string base64="";

            dr.strRtn = "";

            String soqlQuery = "SELECT Id,ParentId,Body FROM Attachment where ParentId='" + Id + "' and Name='" + AttachmentName + "'  order by LastModifiedDate desc limit 1";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);
                Globals.Ribbons.Ribbon1.SFDebug("Get Attachment", soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        base64 = records[i].Any[2].InnerText;
                    }
                }

                //Save as a file
                if (base64 != "")
                {
                    b = Convert.FromBase64String(base64);
                    System.IO.File.WriteAllBytes(filename, b);
                    dr.strRtn = filename;
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }

          
            return dr;
        }

        public DataReturn GetAttachmentFile(string Id,string filename)
        {
            DataReturn dr = new DataReturn();
            byte[] b;
            string base64 = "";

            dr.strRtn = "";

            String soqlQuery = "SELECT Id,ParentId,Body FROM Attachment where Id='" + Id + "'";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);
                Globals.Ribbons.Ribbon1.SFDebug("Get Attachment", soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        base64 = records[i].Any[2].InnerText;
                    }
                }

                //Save as a file
                if (base64 != "")
                {
                    b = Convert.FromBase64String(base64);
                    System.IO.File.WriteAllBytes(filename, b);
                    dr.strRtn = filename;
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }


            return dr;
        }

        public DataReturn SaveAttachmentFile(string ParentId, string AttachmentName, string FileName)
        {
            DataReturn dr = new DataReturn();

            byte[] b;

            try
            {
                b = System.IO.File.ReadAllBytes(FileName);                
            }
            catch (Exception e)
            {
                dr.errormessage = e.Message;
                dr.success = false;
                return dr;
            }

            string id = "";

            String soqlQuery = "SELECT Id FROM Attachment where ParentId='" + ParentId + "' and Name='" + AttachmentName + "'  order by LastModifiedDate desc limit 1";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);
                Globals.Ribbons.Ribbon1.SFDebug("Find Attachment",soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        id = records[i].Any[0].InnerText;
                    }
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }


            sObject attach = new sObject();
            attach.type = "Attachment";
            System.Xml.XmlElement[] o;
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            SaveResult[] sr;

            if (id == "")
            {
                // Create the attacchments fields
                o = new System.Xml.XmlElement[4];
                doc = new System.Xml.XmlDocument();
                o[0] = doc.CreateElement("Name");
                o[0].InnerText = AttachmentName;

                o[1] = doc.CreateElement("isPrivate");
                o[1].InnerText = "false";

                o[2] = doc.CreateElement("ParentId");
                o[2].InnerText = ParentId;

                o[3] = doc.CreateElement("Body");
                o[3].InnerText = Convert.ToBase64String(b);

                attach.Any = o;
                sr = _binding.create(new sObject[] { attach });

                for (int j = 0; j < sr.Length; j++)
                {
                    if (sr[j].success)
                    {
                        id = sr[j].id;
                    }
                    else
                    {
                        for (int i = 0; i < sr[j].errors.Length; i++)
                        {
                            dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];
                        }
                    }
                }
            }
            else
            {
                // Update the attacchments fields
                doc = new System.Xml.XmlDocument();
                o = new System.Xml.XmlElement[1];
                o[0] = doc.CreateElement("Body");
                o[0].InnerText = Convert.ToBase64String(b);

                attach.Any = o;
                attach.Id = id;

                try
                {
                    sr = _binding.update(new sObject[] { attach });
                    Globals.Ribbons.Ribbon1.SFDebug("Update Attachment");


                    for (int j = 0; j < sr.Length; j++)
                    {
                        if (sr[j].success)
                        {
                            id = sr[j].id;
                        }
                        else
                        {
                            for (int i = 0; i < sr[j].errors.Length; i++)
                            {
                                dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    dr.errormessage = e.Message;
                    dr.success = false;
                    return dr;
                }
            }

            dr.id = id;

            return dr;
        }

        public DataReturn UpdateAttachmentFile(string Id,string AttachmentName,string FileName)
        {
            DataReturn dr = new DataReturn();

            byte[] b;

            try
            {
                b = System.IO.File.ReadAllBytes(FileName);
            }
            catch (Exception e)
            {
                dr.errormessage = e.Message;
                dr.success = false;
                return dr;
            }

            sObject attach = new sObject();
            attach.type = "Attachment";
            System.Xml.XmlElement[] o;
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            SaveResult[] sr;

                // Update the attacchments fields
                doc = new System.Xml.XmlDocument();                
                o = new System.Xml.XmlElement[AttachmentName==""?1:2];

                o[0] = doc.CreateElement("Body");
                o[0].InnerText = Convert.ToBase64String(b);

                if (AttachmentName != "")
                {
                    o[1] = doc.CreateElement("Name");
                    o[1].InnerText = AttachmentName;
                }

                attach.Any = o;
                attach.Id = Id;
                try
                {
                    sr = _binding.update(new sObject[] { attach });
                    Globals.Ribbons.Ribbon1.SFDebug("Update Attachment");

                    for (int j = 0; j < sr.Length; j++)
                    {
                        if (sr[j].success)
                        {
                            Id = sr[j].id;
                        }
                        else
                        {
                            for (int i = 0; i < sr[j].errors.Length; i++)
                            {
                                dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i];
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    dr.errormessage = e.Message;
                    dr.success = false;
                    return dr;
                }
            

            dr.id = Id;

            return dr;
        }

        public DataReturn ExecRibbonCall(string Action,string Obj,string Id)
        {
            DataReturn dr = new DataReturn();

            sfRibbon.RibbonRequest req = new RibbonRequest();
            req.action = Action;
            req.objname = Obj;
            req.id = Id;

            try
            {
                sfRibbon.RibbonResponse res = _ribbonBinding.Dispatch(req);

                Globals.Ribbons.Ribbon1.SFDebug("Call Axiom_RibbonCntroller", "Action:" + Action + " Obj:" + Obj + " Id:" + Id);

                dr.success = (res.success == null ? false : (res.success == true ? true : false));
                dr.id = res.selectid;
                dr.reload = res.reload;
                dr.strRtn = res.message;
            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.strRtn = ex.Message;
            }

            return dr;

        }




        string CleanUpXML(string val)
        {
            string rtn = Regex.Replace(val, @"[\x00-\x08]|[\x0B\x0C]|[\x0E-\x19]|[\uD800-\uDFFF]|[\uFFFE\uFFFF]", "");
            rtn = Regex.Replace(rtn, @"[\x1A-\x1F]", "-");
            return rtn;
        }

        /* ---------------------------------------------------------Generic SForce Routines - need to clean up and merge with above */

        //Add Data to the Dataset - need to call recursilvely
        private void AddData(SForceEdit.SObjectDef sobj, DataRow rw, System.Xml.XmlElement x, string ParentName)
        {
            if (ParentName != "") ParentName += "_";

            if (x.HasAttributes && x.Attributes["xsi:type"] != null && x.Attributes["xsi:type"].Value == "sf:sObject")
            {
                string tempname = ParentName + x.LocalName;
                for (int k = 0; k < x.ChildNodes.Count; k++)
                {
                    System.Xml.XmlElement xchild = ((System.Xml.XmlElement)x.ChildNodes[k]);
                    if (xchild.HasAttributes && xchild.Attributes["xsi:type"] != null && xchild.Attributes["xsi:type"].Value == "sf:sObject")
                    {
                        AddData(sobj, rw, xchild, tempname);
                    }
                    else
                    {
                        //Check if the table contains the field if it does add it in 
                        string fldname = tempname + "_" + xchild.LocalName;
                        if (rw.Table.Columns.Contains(fldname))
                        {
                            rw[fldname] = GetDataColumnData(fldname, xchild.InnerText);
                        }
                        else
                        {
                            //if the column is not there add it in if its in the columns objects
                            if (sobj.FieldExists(fldname))
                            {
                                string fldtype = sobj.GetField(fldname).DataType;
                                System.Data.DataColumn c = new DataColumn(fldname, GetDataColumnType(fldtype));
                                rw.Table.Columns.Add(c);
                                rw[fldname] = GetDataColumnData(fldname, xchild.InnerText);
                            }
                        }
                    }
                }
            }
            else
            {
                if (rw.Table.Columns.Contains(x.LocalName))
                {
                    string fldtype = sobj.GetField(x.LocalName).DataType;
                    rw[x.LocalName] = GetDataColumnData(fldtype, x.InnerText);
                }
                else
                {
                    //if the column is not there add it in if its in the columns objects
                    if (sobj.FieldExists(x.LocalName))
                    {
                        string fldtype = sobj.GetField(x.LocalName).DataType;
                        System.Data.DataColumn c = new DataColumn(x.LocalName, GetDataColumnType(fldtype));
                        rw.Table.Columns.Add(c);
                        rw[x.LocalName] = GetDataColumnData(fldtype, x.InnerText);
                    }
                }
            }

        }

        //Given the SForce DataType return the data co,um type
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

        //Bit of data checking - can't set dates or doubles to "" - so change to null
        private object GetDataColumnData(string DataType, string val)
        {
            if (val == "")
            {
                if (DataType == "date" || DataType == "datetime")
                {
                    return DBNull.Value;
                }
                else if (DataType == "double" || DataType == "currency")
                {
                    return DBNull.Value;
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return val;
            }
        }

        public DataReturn RunSOQL(SForceEdit.SObjectDef sObj)
        {

            DataReturn dr = new DataReturn();

            //Build Query
            string soqlQuery = "select " + sObj.GetQueryList() + " from " + sObj.Name + "";

            //See if there is a fitler
            string filter = "";
            string filterdefaultsort = "";

            // see if we have defined filters
            if (sObj.GridFilters.Count > 0)
            {
                // defined filters
                // get the userid - check if there is an error, this could be the first thing to be called after a timeout
                string userid = "";
                try
                {
                    userid = GetUserId();
                }
                catch (Exception ex)
                {
                    dr.success = false;
                    dr.errormessage = ex.Message;
                    return dr;
                }


                if (sObj.Filter!=null && sObj.GridFilters.ContainsKey(sObj.Filter))
                {
                    AxiomIRISRibbon.SForceEdit.SObjectDef.FilterEntry f = sObj.GridFilters[sObj.Filter];
                    string soql = f.SOQL;
                    soql = soql.Replace("{UserId}", userid);

                    if(soql!="") filter = " where " + soql;
                    filterdefaultsort = f.OrderBy==""?"":" ORDER BY " + f.OrderBy;
                }

            }
           
            if (sObj.Search != "")
            {
                filter += (filter == "" ? " where " : " and ");
                filter += sObj.GetSearchClause();
            }

            if (sObj.Parent != "")
            {
                filter += (filter == "" ? " where " : " and ");
                string pclause = sObj.Parent;
                if (!pclause.EndsWith("__c"))
                {
                    pclause += "Id";
                }

                // if the parent id is blank we really want none - this is when we create a new parent and the 
                // subtab is selected - can't just be parent='' cause there maybe children that have no parent and SOQL
                // returns them if you do ='' - so make it a dummy Id instead
                if (sObj.ParentId == "")
                {
                    filter += pclause + " = '123456789012345678'";
                }
                else
                {
                    filter += pclause + " = '" + sObj.ParentId + "'";
                }
            }

            if (sObj.Id != "")
            {
                filter += (filter == "" ? " where " : " and ");
                filter += "Id = '" + sObj.Id + "'";
            }

            soqlQuery += filter;

            if (sObj.Paging)
            {
                if (sObj.SortColumn != "")
                {
                    soqlQuery += " ORDER BY " + sObj.SortQueryField + " " + sObj.SortDir;
                }
                else
                {
                    soqlQuery += filterdefaultsort;                    
                }

                soqlQuery += " LIMIT " + sObj.RecordsPerPage.ToString() + " OFFSET " + (sObj.CurrnetPage * sObj.RecordsPerPage).ToString();

                try
                {
                    //Get total count
                    QueryResult qr = _binding.query("select count() from " + sObj.Name + filter);
                    Globals.Ribbons.Ribbon1.SFDebug("GetCount>" + sObj.Name, "select count() from " + sObj.Name + filter);
                    sObj.TotalRecords = qr.size;
                }
                catch (Exception ex)
                {
                    dr.success = false;
                    dr.errormessage = ex.Message;
                    return dr;
                }

            }
            else
            {
                if (sObj.SortColumn == null || sObj.SortColumn == "")
                {
                    soqlQuery += filterdefaultsort;
                }
            }
            

            
            
            //Create the DataTable from the Definition
            dr.dt = sObj.CreateDataTable();

            try
            {


                QueryResult qr = _binding.query(soqlQuery);
                Globals.Ribbons.Ribbon1.SFDebug("Get>" + sObj.Name, soqlQuery);
                bool done = false;

                if (qr.size > 0)
                {
                    while (!done)
                    {
                        sObject[] records = qr.records;
                        for (int i = 0; i < qr.records.Length; i++)
                        {

                            DataRow rw = dr.dt.NewRow();
                            for (int j = 0; j < records[i].Any.Length; j++)
                            {
                                AddData(sObj, rw, records[i].Any[j], "");
                            }

                            dr.dt.Rows.Add(rw);

                        }

                        if (qr.done)
                        {
                            done = true;
                        }
                        else
                        {
                            qr = _binding.queryMore(qr.queryLocator);
                            Globals.Ribbons.Ribbon1.SFDebug("GetMore>" + sObj.Name, "More>" + soqlQuery);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }

            return dr;
        }

        //Given a DataRow, update or Create the SalesForce Object
        //Assuming that we have just one row, easy to change to handle multiples
        public DataReturn Save(SForceEdit.SObjectDef sObj, DataRow dRow)
        {

            DataReturn dr = new DataReturn();

            sObject s = new sObject();
            s.type = sObj.Name;
            string id = "";

            if (dRow["Id"] == null || dRow["Id"].ToString() == "")
            {
                //new
                int fldCount = dRow.Table.Columns.Count - 1;
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                System.Xml.XmlElement[] o = new System.Xml.XmlElement[fldCount];

                fldCount = 0;

                List<string> fieldsToNull = new List<string>();

                foreach (DataColumn dc in dRow.Table.Columns)
                {

                    //Get the field definition
                    SForceEdit.SObjectDef.FieldGridCol f = sObj.GetField(dc.ColumnName);

                    // this is a new record so do it even if it says its readonly but exclud any _Name or _Type
                    if (!f.Create)
                    {
                        //nada ...
                    }
                    else if (dc.ColumnName == "Id")
                    {
                        //Nothing!
                    }
                    else if (dc.ColumnName == "type")
                    {
                        //don't do anything - this happens when we have the type field from a join
                    }
                    else
                    {

                        object val = dRow[dc.ColumnName];
                        if (dRow[dc.ColumnName] == DBNull.Value)
                        {
                            fieldsToNull.Add(dc.ColumnName);
                        }
                        else
                        {
                            o[fldCount] = doc.CreateElement(dc.ColumnName);

                            string sval = "";
                            if (f.DataType == "datetime")
                            {
                                sval = ((DateTime)val).ToString("o");
                            }
                            else if (f.DataType == "date")
                            {
                                sval = ((DateTime)val).ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                sval = CleanUpXML(val.ToString());
                            }

                            o[fldCount].InnerText = sval;
                            fldCount++;
                        }

                    }
                }

                try
                {
                    // dont need to set the values to Null! this is a create so just don't tell them
                    // s.fieldsToNull = fieldsToNull.ToArray();
                    s.Any = Utility.SubArray<System.Xml.XmlElement>(o, 0, fldCount);
                    sfPartner.SaveResult[] sr = _binding.create(new sObject[] { s });
                    Globals.Ribbons.Ribbon1.SFDebug("Save>" + s.type);

                    for (int j = 0; j < sr.Length; j++)
                    {
                        if (sr[j].success)
                        {
                            dr.id = sr[j].id;
                        }
                        else
                        {
                            dr.success = false;
                            for (int i = 0; i < sr[j].errors.Length; i++)
                            {
                                dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i].message;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    dr.success = false;
                    dr.errormessage = ex.Message;
                }
            }
            else
            {
                //update
                int fldCount = dRow.Table.Columns.Count;
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();

                System.Xml.XmlElement[] o = new System.Xml.XmlElement[fldCount];

                fldCount = 0;

                List<string> fieldsToNull = new List<string>();

                foreach (DataColumn dc in dRow.Table.Columns)
                {
                    //Get the field definition
                    SForceEdit.SObjectDef.FieldGridCol f = sObj.GetField(dc.ColumnName);

                    if (dc.ColumnName == "Id")
                    {
                        s.Id = dRow[dc.ColumnName].ToString();
                    }
                    else if (!f.Update)
                    {
                        //not on the list ...
                    }
                    else if (dc.ColumnName == "type")
                    {
                        //don't do anything - this happens when we have the type field from a join
                    }
                    else
                    {

                        object val = dRow[dc.ColumnName];
                        if (dRow[dc.ColumnName] == DBNull.Value ||
                            ((f.DataType != "string") && dRow[dc.ColumnName].ToString() == ""))
                        {
                            fieldsToNull.Add(dc.ColumnName);
                        }
                        else
                        {
                            o[fldCount] = doc.CreateElement(dc.ColumnName);

                            string sval = "";
                            if (f.DataType == "datetime")
                            {
                                sval = ((DateTime)val).ToString("o");
                            }
                            else if (f.DataType == "date")
                            {
                                sval = ((DateTime)val).ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                sval = CleanUpXML(val.ToString());
                            }


                            o[fldCount].InnerText = sval;
                            fldCount++;
                        }
                    }
                }

                try
                {
                    s.fieldsToNull = fieldsToNull.ToArray();
                    s.Any = Utility.SubArray<System.Xml.XmlElement>(o, 0, fldCount);
                    sfPartner.SaveResult[] sr = _binding.update(new sObject[] { s });
                    Globals.Ribbons.Ribbon1.SFDebug("Update>" + s.type);
                    for (int j = 0; j < sr.Length; j++)
                    {
                        Console.WriteLine("\nItem: " + j);
                        if (sr[j].success)
                        {
                            dr.id = sr[j].id;
                        }
                        else
                        {
                            dr.success = false;
                            for (int i = 0; i < sr[j].errors.Length; i++)
                            {
                                dr.errormessage += (dr.errormessage == "" ? "" : ",") + sr[j].errors[i].message;
                            }
                        }
                    }


                }
                catch (Exception ex)
                {
                    dr.success = false;
                    dr.errormessage = ex.Message;
                }

            }
            return dr;
        }




        public DataReturn LoadDefinitions(string[] sObjects)
        {

            DataReturn dr = new DataReturn();
            _allSObjects = new Dictionary<string, DescribeGlobalSObjectResult>();
            _describeSObject = new Dictionary<string, DescribeSObjectResult>();
            _describeSearch = new Dictionary<string, DescribeSearchLayoutResult>();
            _describeLayout = new Dictionary<string, DescribeLayoutResult>();

            try
            {
                //First get the Global List of objects
                DescribeGlobalResult dgr = _binding.describeGlobal();
                Globals.Ribbons.Ribbon1.SFDebug("Global Describe");
                DescribeGlobalSObjectResult[] sObjResults = dgr.sobjects;
                
                
                for (int i = 0; i < sObjResults.Length; i++)
                {
                    _allSObjects.Add(sObjResults[i].name, sObjResults[i]);
                }

                //Check the defined objects exist
                List<string> sObjectsExist = new List<string>();
                for (int i = 0; i < sObjects.Length; i++)
                {
                    if(_allSObjects.ContainsKey(sObjects[i])) sObjectsExist.Add(sObjects[i]);
                }
                sObjects = sObjectsExist.ToArray();
                

                DescribeSObjectResult[] dso = _binding.describeSObjects(sObjects);
                Globals.Ribbons.Ribbon1.SFDebug("Describe Objects"+ string.Join("|",sObjects));
                for (int i = 0; i < dso.Length; i++)
                {
                    _describeSObject.Add(sObjects[i], dso[i]);
                }

                /* this doesn't actually get what we need! it gives what the search returns in SForce when you do a generic accross object saerch
                   wnated to get the ListView but can only get using MetaData and have to be admin - so using config for now
                // Think this is for search pages - remove Task, it doesn't have a search layout
                string[] sObjectsWithoutTasks = sObjects.Where(val => val != "Task").ToArray();
                DescribeSearchLayoutResult[] dslr = _binding.describeSearchLayouts(sObjectsWithoutTasks);
                Globals.Ribbons.Ribbon1.SFDebug("Describe Layouts" + string.Join("|", sObjectsWithoutTasks));
                for (int i = 0; i < dslr.Length; i++)
                {
                    _describeSearch.Add(sObjectsWithoutTasks[i], dslr[i]);
                }
                */

                for (int i = 0; i < sObjects.Length; i++)
                {
                    //Attachment doesn't have a layout
                    if (sObjects[i] != "Attachment")
                    {
                        _describeLayout.Add(sObjects[i], _binding.describeLayout(sObjects[i], null,null));
                        Globals.Ribbons.Ribbon1.SFDebug("Describe Layouts for " + sObjects[i]);
                    }
                }

            }
            catch (Exception e)
            {
                dr.success = false;
                dr.errormessage = e.Message;
            }
            return dr;
        }

        public DescribeLayoutResult GetDefinitionLayout(string name)
        {
            return _describeLayout.ContainsKey(name)?_describeLayout[name]:null;
        }

        public DescribeSObjectResult GetDefinitionSObject(string name)
        {
            return _describeSObject.ContainsKey(name) ? _describeSObject[name] : null;
        }

        public DescribeSearchLayoutResult GetDefinitionSearch(string name)
        {
            if (_describeSearch.ContainsKey(name))
            {
                return _describeSearch[name];
            }
            else
            {
                return null;
            }
        }

        public DescribeGlobalSObjectResult GetSObjectDef(string name)
        {
            return _allSObjects[name];
        }


        public string GetUrlForNonLoaded(string name)
        {
            DescribeSObjectResult dso = _binding.describeSObject(name);
            return dso.urlDetail;
        }

        // get the specified static file as a string
        public DataReturn GetStaticResource(string name)
        {
            DataReturn dr = new DataReturn();
            byte[] b;
            string base64 = "";

            dr.strRtn = "";

            String soqlQuery = "SELECT Id,Body FROM StaticResource where Name='" + name + "'";
            try
            {
                QueryResult qr = _binding.query(soqlQuery);
                Globals.Ribbons.Ribbon1.SFDebug("Get Settings File", soqlQuery);

                if (qr.size > 0)
                {
                    sObject[] records = qr.records;
                    for (int i = 0; i < qr.records.Length; i++)
                    {
                        base64 = records[i].Any[1].InnerText;
                    }
                }

                //Convert to string
                if (base64 != "")
                {
                    b = Convert.FromBase64String(base64);
                    dr.strRtn = Encoding.UTF8.GetString(b);
                }

            }
            catch (Exception ex)
            {
                dr.success = false;
                dr.errormessage = ex.Message;
            }


            return dr;
        }

    }
}
