using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace AxiomIRISRibbon.SForceEdit
{

    // TODO Switch to using JSON For the Settings at some point! all these split strings are horrible!

    public class Settings
    {
        JObject RibbonSettings;

        public Settings(string settings)
        {
            this.RibbonSettings = null;
            this.RibbonSettings = JObject.Parse(settings);
                       
        }

        public void AddSetting(string sObject, string key, string val)
        {
            
        }


        public JToken GetGeneralSetting(string key)
        {
            if (this.RibbonSettings["General"] != null)
            {
                return this.RibbonSettings["General"][key];
            }
            else
            {
                return null;
            }
        }

        public string GetSetting(string sObject, string key)
        {
            return GetSetting(sObject, key, "");
        }

        public string GetSetting(string sObject, string key,string subkey)
        {
            string rtn = "";
            if (this.RibbonSettings != null)
            {
                if (key == "SObjects")
                {
                    foreach (JObject sobj in this.RibbonSettings["SForceObjects"])
                    {
                        if (sobj["Name"] != null)
                        {
                            rtn += (rtn == "" ? "" : "|") + sobj["Name"];
                        }
                    }
                }
                else if (key == "TopLevelSObjects")
                {

                    foreach (JObject sobj in this.RibbonSettings["SForceObjects"])
                    {
                        JToken TopLevel = sobj["TopLevel"];
                        if (sobj["TopLevel"] != null && Convert.ToBoolean(sobj["TopLevel"]) == true)
                        {
                            if (sobj["Name"] != null)
                            {
                                string label = (sobj["Label"] == null ? "" : sobj["Label"].ToString());
                                string icon = (sobj["Icon"] == null ? "" : sobj["Icon"].ToString());

                                if (label == "") label = sobj["Name"].ToString().Replace("__c","").Replace("_", " ");

                                rtn += (rtn == "" ? "" : "|") + sobj["Name"] + ":" + label + ":" + icon;
                            }                            
                        }

                    }

                }
                else
                {
                    foreach (JObject sobj in this.RibbonSettings["SForceObjects"])
                    {
                        if (sobj["Name"].ToString() == sObject)
                        {
                            if (key == "Tabs")
                            {
                                if (sobj["Tabs"] != null)
                                {
                                    foreach (JObject stab in sobj["Tabs"])
                                    {
                                        rtn += (rtn == "" ? "" : "|") + stab["SubObject"] + ":" + stab["ParentRelationName"];
                                    }
                                }
                            }
                            else if (key == "Columns")
                            {
                                if (sobj["Columns"] != null){
                                foreach (string scol in sobj["Columns"])
                                {
                                    rtn += (rtn == "" ? "" : "|") + scol;
                                }
                            }
                            }
                            else if (key == "Buttons")
                            {
                                if (sobj["Buttons"] != null)
                                {
                                    foreach (var sbutton in sobj["Buttons"])
                                    {
                                        if (sbutton.Type == JTokenType.Object)
                                        {
                                            // TODO for now return as string but should really pass a
                                            // dynamic object or something or just JSON
                                            string btn = "";
                                            if (sbutton["Name"] != null || sbutton["Type"] != null)
                                            {
                                                btn = sbutton["Name"] == null ? "" : sbutton["Name"].ToString();
                                                btn += ":" + (sbutton["Type"] == null ? "Data":sbutton["Type"].ToString());
                                                btn += ":" + (sbutton["Action"] == null ? "" : sbutton["Action"].ToString());

                                                string rtypes = "";
                                                if (sbutton["RecordTypes"] != null)
                                                {
                                                    if (sbutton["RecordTypes"].Type == JTokenType.Array)
                                                    {
                                                        foreach (var rt in sbutton["RecordTypes"])
                                                        {
                                                            rtypes += (rtypes == "" ? "" : ",") + rt;
                                                        }
                                                    }
                                                    else if (sbutton["RecordTypes"].Type == JTokenType.String)
                                                    {
                                                        rtypes += sbutton["RecordTypes"].ToString();
                                                    }
                                                }
                                                btn += ":" + rtypes;

                                                string confirm = "";
                                                if (sbutton["Confirm"] != null)
                                                {
                                                    confirm = sbutton["Confirm"].ToString();
                                                }
                                                btn += ":" + confirm;

                                                rtn += (rtn == "" ? "" : "|") + btn;
                                            }
                                        }
                                    }
                                }
                            }
                            else if (key == "Compact")
                            {
                                if (sobj["Compact"] != null){
                                    if (sobj["Compact"].Type == JTokenType.Array)
                                    {
                                        foreach (string scom in sobj["Compact"])
                                        {
                                            rtn += (rtn == "" ? "" : "|") + scom;
                                        }
                                    }
                                    else if (sobj["Compact"].Type == JTokenType.String)
                                    {
                                        rtn = sobj["Compact"].ToString();
                                    }
                            }
                            }
                            else if (key == "Filters")
                            {
                                if (sobj["Filters"] != null)
                                {
                                    if (sobj["Filters"].Type == JTokenType.Array)
                                    {
                                        foreach (JObject f in sobj["Filters"])
                                        {
                                            string name = f["Name"] == null ? "" : f["Name"].ToString();
                                            string soql = f["SOQL"] == null ? "" : f["SOQL"].ToString();
                                            string def = f["Default"] == null ? "" : f["Default"].ToString();
                                            string orderby = f["OrderBy"] == null ? "" : f["OrderBy"].ToString();
                                            rtn += (rtn == "" ? "" : "|") + name + ":" + soql + ":" + def + ":" + orderby;
                                        }
                                    }                                    
                                }
                            }
                            else if (key == "TabFilters")
                            {
                                if (sobj["Tabs"] != null)
                                {
                                    foreach (JObject stab in sobj["Tabs"])
                                    {
                                        if (stab["SubObject"].ToString() == subkey)
                                        {
                                            if (stab["Filters"] != null)
                                            {
                                                if (stab["Filters"].Type == JTokenType.Array)
                                                {
                                                    foreach (JObject f in stab["Filters"])
                                                    {
                                                        string name = f["Name"] == null ? "" : f["Name"].ToString();
                                                        string soql = f["SOQL"] == null ? "" : f["SOQL"].ToString();
                                                        string def = f["Default"] == null ? "" : f["Default"].ToString();
                                                        string orderby = f["OrderBy"] == null ? "" : f["OrderBy"].ToString();
                                                        rtn += (rtn == "" ? "" : "|") + name + ":" + soql + ":" + def + ":" + orderby;
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
            return rtn;
        }
    }

    

}
