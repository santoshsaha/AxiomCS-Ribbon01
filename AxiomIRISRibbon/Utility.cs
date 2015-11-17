using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Threading;
using System.Security.Permissions;
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
using System.Collections;
using System.Security.Cryptography;
using System.IO;
using System.Text.RegularExpressions;

namespace AxiomIRISRibbon
{
    public static class Utility
    {

        //Add a DoEvents - shoudln't really do it like this but I'm lazy
        [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        public static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
                new DispatcherOperationCallback(ExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        public static object ExitFrame(object f)
        {
            ((DispatcherFrame)f).Continue = false;

            return null;
        }


        // Handy form functions - pass in a DataTable and the form and it will update any 
        // matching fields - to pass in the form pass all the Grid parents with form fields inside
        // did try to use the logical/visual trees but had lots of issues! visual tree only shows visible and logical didn't seem to work
        static public void UpdateForm(Grid[] grds, DataRow dr)
        {

            foreach (Grid g in grds)
            {
                foreach (object child in g.Children)
                {
                    Control childControl = child as Control;

                    if (childControl != null)
                    {
                        //if there is a scroll then get the child element
                        if (childControl.GetType() == typeof(ScrollViewer))
                        {
                            ScrollViewer sc = (ScrollViewer)childControl;
                            if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox))
                            {
                                childControl = (Control)sc.Content;
                            }
                        }

                        if (childControl.GetType() == typeof(TextBox))
                        {
                            TextBox tb = (TextBox)childControl;
                            if (tb.Tag == null || tb.Tag.ToString() != "ignore")
                            {
                                string name = tb.Name;
                                //bit of mungling to get the right name
                                if (name.StartsWith("tb")) name = name.Substring(2, name.Length - 2);

                                //Allow overite with the tag field
                                if (tb.Tag != null && tb.Tag.ToString() != "")
                                {
                                    //text field can be set to readonly so ignore anything after |
                                    string[] flddef = tb.Tag.ToString().Split('|');
                                    name = flddef[0];
                                }

                                string val = "";

                                //Check the datatable for that column
                                foreach (DataColumn dc in dr.Table.Columns)
                                {
                                    if (dc.ColumnName == name || dc.ColumnName == name + "__c") val = dr[dc.ColumnName].ToString();
                                }

                                tb.Text = val;
                            }
                        }
                        else if (childControl.GetType() == typeof(ComboBox))
                        {
                            ComboBox cb = (ComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                string idfield = "";
                                // bit of mungling to get the right name
                                if (name.StartsWith("cb")) name = name.Substring(2, name.Length - 2);

                                // Allow overite with the tag field
                                // For combo this defines the list field and the id field
                                if (cb.Tag != null && cb.Tag.ToString() != "")
                                {
                                    //DisplayField|IdField
                                    string[] flddef = cb.Tag.ToString().Split('|');
                                    name = flddef[0];
                                    if (flddef.Length > 1) idfield = flddef[1];
                                }

                                string val = "";

                                // Russel 2 Oct - Oh my - this is selecting on Name not id - if there are duplicate names this can 
                                // be an issue! fix to work on the id if there is one!

                                if (idfield == "")
                                {
                                    //Check the datatable for that column
                                    foreach (DataColumn dc in dr.Table.Columns)
                                    {
                                        if (dc.ColumnName == name || dc.ColumnName == name + "__c") val = dr[dc.ColumnName].ToString();
                                    }

                                    cb.Text = val;
                                }
                                else
                                {
                                    foreach (DataColumn dc in dr.Table.Columns)
                                    {
                                        if (dc.ColumnName == idfield || dc.ColumnName == idfield + "__c") val = dr[dc.ColumnName].ToString();
                                    }

                                    ComboBoxItem selectitem = null;
                                    foreach (var item in cb.Items)
                                    {
                                        ComboBoxItem cbi = (ComboBoxItem)item;
                                        if(cbi.Tag!=null){
                                            if (cbi.Tag.ToString() == val)
                                            {
                                                selectitem = cbi;
                                            }
                                        }
                                        
                                    }
                                    cb.SelectedItem = selectitem;                                    
                                }
                            }
                        }

                        else if (childControl.GetType() == typeof(RadComboBox))
                        {
                            RadComboBox cb = (RadComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                string idfield = "";
                                //bit of mungling to get the right name
                                if (name.StartsWith("cb")) name = name.Substring(2, name.Length - 2);


                                //Allow overite with the tag field
                                //For combo this defines the list field and the id field
                                if (cb.Tag != null && cb.Tag.ToString() != "")
                                {
                                    //DisplayField|IdField
                                    string[] flddef = cb.Tag.ToString().Split('|');
                                    name = flddef[0];
                                    if (flddef.Length > 1) idfield = flddef[1];
                                }


                                string val = "";

                                // Russel 2 Oct - Oh my - this is selecting on Name not id - if there are duplicate names this can 
                                // be an issue!

                                if (idfield == "")
                                {
                                    //Check the datatable for that column
                                    foreach (DataColumn dc in dr.Table.Columns)
                                    {
                                        if (dc.ColumnName == name || dc.ColumnName == name + "__c") val = dr[dc.ColumnName].ToString();
                                    }

                                    cb.Text = val;
                                }
                                else
                                {
                                    foreach (DataColumn dc in dr.Table.Columns)
                                    {
                                        if (dc.ColumnName == idfield || dc.ColumnName == idfield + "__c") val = dr[dc.ColumnName].ToString();
                                    }

                                    RadComboBoxItem selectitem = null;
                                    foreach (var item in cb.Items)
                                    {
                                        RadComboBoxItem cbi = (RadComboBoxItem)item;
                                        if (cbi.Tag != null)
                                        {
                                            if (cbi.Tag.ToString() == val)
                                            {
                                                selectitem = cbi;
                                            }
                                        }

                                    }
                                    cb.SelectedItem = selectitem;
                                }
                            }
                        }
                        else if (childControl.GetType() == typeof(CheckBox))
                        {
                            CheckBox cb = (CheckBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                //bit of mungling to get the right name
                                if (name.StartsWith("cb")) name = name.Substring(2, name.Length - 2);

                                //Allow overite with the tag field
                                if (cb.Tag != null && cb.Tag.ToString() != "")
                                {
                                    name = cb.Tag.ToString();
                                }

                                string val = "";

                                //Check the datatable for that column
                                foreach (DataColumn dc in dr.Table.Columns)
                                {
                                    if (dc.ColumnName == name || dc.ColumnName == name + "__c") val = dr[dc.ColumnName].ToString();
                                }

                                if (val == "") val = "false";
                                if (Convert.ToBoolean(val))
                                {
                                    cb.IsChecked = true;
                                }
                                else
                                {
                                    cb.IsChecked = false;
                                }

                            }
                        }
                    }
                }
            }

        }
        

        //Other way round - update the datarow from the form
        static public void UpdateRow(Grid[] grds, DataRow dr)
        {
            foreach (Grid g in grds)
            {
                foreach (object child in g.Children)
                {
                    Control childControl = child as Control;

                    if (childControl != null)
                    {
                        //if there is a scroll then get the child element
                        if (childControl.GetType() == typeof(ScrollViewer))
                        {
                            ScrollViewer sc = (ScrollViewer)childControl;
                            if(sc.Content.GetType()  == typeof(TextBox) || sc.Content.GetType()  == typeof(ComboBox)){
                                childControl = (Control)sc.Content;
                            }
                        }

                        if (childControl.GetType() == typeof(TextBox))
                        {
                            TextBox tb = (TextBox)childControl;

                            if (tb.Tag == null || tb.Tag.ToString() != "ignore")
                            {
                                string name = tb.Name;
                                if (name.StartsWith("tb")) name = name.Substring(2, name.Length - 2);

                                //Allow overite with the tag field
                                string idfield = "";
                                if (tb.Tag != null && tb.Tag.ToString() != "")
                                {
                                    //text field can be set to readonly so ignore anything after |
                                    string[] flddef = tb.Tag.ToString().Split('|');
                                    name = flddef[0];

                                }

                                string val = "";

                                //Check the datatable for that column
                                foreach (DataColumn dc in dr.Table.Columns)
                                {
                                    if (dc.ColumnName == name || dc.ColumnName == name + "__c")
                                    {
                                        val = tb.Text;
                                        dr[dc.ColumnName] = val;
                                    }
                                }
                            }
                        }
                        else if (childControl.GetType() == typeof(ComboBox))
                        {
                            ComboBox cb = (ComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                if (name.StartsWith("cb")) name = name.Substring(2, name.Length - 2);

                                //Allow overite with the tag field
                                //For combo this defines the list field and the id field
                                //the id field is set in the tag field of the combo list - need to update both
                                string idfield = "";
                                if(cb.Tag!=null && cb.Tag.ToString() != "")
                                {
                                    //DisplayField|IdField
                                    string[] flddef = cb.Tag.ToString().Split('|');
                                    name = flddef[0];
                                    if (flddef.Length > 1) idfield = flddef[1];
                                }

                                string val = "";

                                //Check the datatable for that column
                                foreach (DataColumn dc in dr.Table.Columns)
                                {
                                    if (dc.ColumnName == name || dc.ColumnName == name + "__c")
                                    {
                                        val = cb.Text;
                                        dr[dc.ColumnName] = val;
                                    }
                                    if (dc.ColumnName == idfield)
                                    {
                                        val = ((ComboBoxItem)cb.SelectedItem).Tag.ToString();
                                        dr[dc.ColumnName] = val;
                                    }

                                }
                            }
                        }
                        else if (childControl.GetType() == typeof(RadComboBox))
                        {
                            RadComboBox cb = (RadComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                if (name.StartsWith("cb")) name = name.Substring(2, name.Length - 2);

                                //Allow overite with the tag field
                                //For combo this defines the list field and the id field
                                //the id field is set in the tag field of the combo list - need to update both
                                string idfield = "";
                                if (cb.Tag != null && cb.Tag.ToString() != "")
                                {
                                    //DisplayField|IdField
                                    string[] flddef = cb.Tag.ToString().Split('|');
                                    name = flddef[0];
                                    if (flddef.Length > 1) idfield = flddef[1];
                                }

                                string val = "";

                                //Check the datatable for that column
                                foreach (DataColumn dc in dr.Table.Columns)
                                {
                                    if (dc.ColumnName == name || dc.ColumnName == name + "__c")
                                    {
                                        val = cb.Text;
                                        dr[dc.ColumnName] = val;
                                    }
                                    if (dc.ColumnName == idfield)
                                    {
                                        if (cb.SelectedItem != null)
                                        {
                                            val = ((RadComboBoxItem)cb.SelectedItem).Tag.ToString();
                                            dr[dc.ColumnName] = val;
                                        }
                                    }

                                }
                            }
                        }
                        else if (childControl.GetType() == typeof(CheckBox))
                        {
                            CheckBox cb = (CheckBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                //bit of mungling to get the right name
                                if (name.StartsWith("cb")) name = name.Substring(2, name.Length - 2);

                                //Allow overite with the tag field
                                if (cb.Tag != null && cb.Tag.ToString() != "")
                                {
                                    name = cb.Tag.ToString();
                                }

                                string val = "";

                                //Check the datatable for that column
                                foreach (DataColumn dc in dr.Table.Columns)
                                {
                                    if (dc.ColumnName == name || dc.ColumnName == name + "__c")
                                    {
                                        val = cb.IsChecked.ToString();
                                        dr[dc.ColumnName] = val;
                                    }
                                }

                            }
                        }
                    }
                }
            }

        }

        static public void ClearForm(Grid[] grds)
        {

            foreach (Grid g in grds)
            {
                foreach (object child in g.Children)
                {
                    Control childControl = child as Control;

                    if (childControl != null)
                    {
                        //if there is a scroll then get the child element
                        if (childControl.GetType() == typeof(ScrollViewer))
                        {
                            ScrollViewer sc = (ScrollViewer)childControl;
                            if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox))
                            {
                                childControl = (Control)sc.Content;
                            }
                        }

                        if (childControl.GetType() == typeof(TextBox))
                        {
                            TextBox tb = (TextBox)childControl;
                            if (tb.Tag == null || tb.Tag.ToString() != "ignore")
                            {
                                string name = tb.Name;
                                tb.Text = "";
                            }
                        }
                        else if (childControl.GetType() == typeof(ComboBox))
                        {
                            ComboBox cb = (ComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                cb.Text = "";
                            }
                        }
                        else if (childControl.GetType() == typeof(RadComboBox))
                        {
                            RadComboBox cb = (RadComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                string name = cb.Name;
                                cb.Text = "";
                            }
                        }
                        else if (childControl.GetType() == typeof(CheckBox))
                        {
                            CheckBox cb = (CheckBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                cb.IsChecked = false;
                            }
                        }
                    }
                }
            }

        }


        static public void ReadOnlyForm(bool IsReadOnly,Grid[] grds)
        {

            foreach (Grid g in grds)
            {
                foreach (object child in g.Children)
                {
                    Control childControl = child as Control;

                    if (childControl != null)
                    {
                        //if there is a scroll then get the child element
                        if (childControl.GetType() == typeof(ScrollViewer))
                        {
                            ScrollViewer sc = (ScrollViewer)childControl;
                            if (sc.Content.GetType() == typeof(TextBox) || sc.Content.GetType() == typeof(ComboBox))
                            {
                                childControl = (Control)sc.Content;
                            }
                        }

                        if (childControl.GetType() == typeof(TextBox))
                        {
                            TextBox tb = (TextBox)childControl;
                            if (tb.Tag == null || tb.Tag.ToString() != "ignore")
                            {
                                bool update = true;
                                if (tb.Tag != null)
                                {
                                    if (tb.Tag.ToString().Trim().ToLower().EndsWith("|readonly"))
                                    {
                                        update = false;
                                    }
                                }

                                if(update) tb.IsReadOnly = IsReadOnly;
                            }
                        }
                        else if (childControl.GetType() == typeof(ComboBox))
                        {
                            ComboBox cb = (ComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                cb.IsEnabled = !IsReadOnly;
                            }
                        }
                        else if (childControl.GetType() == typeof(RadComboBox))
                        {
                            RadComboBox cb = (RadComboBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                cb.IsEnabled = !IsReadOnly;
                            }
                        }
                        else if (childControl.GetType() == typeof(CheckBox))
                        {
                            CheckBox cb = (CheckBox)childControl;
                            if (cb.Tag == null || cb.Tag.ToString() != "ignore")
                            {
                                cb.IsEnabled = !IsReadOnly;
                            }
                        }
                    }
                }
            }

        }

        public static T[] SubArray<T>(this T[] data, int index, int length)
        {
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }

        public static Boolean StyleExists(Word.Styles styles,string name)
        {
            foreach (Word.Style s in styles)
            {
                if (s.NameLocal == name) return true;
            }
            return false;
        }


        public static string SaveTempFile(string id){
             // Generate a temp file and save the doc there

             // check the id isn't actually a filename
            if (id.ToLower().EndsWith(".docx"))
            {
                id = id.Substring(0, id.Length - 5);
            }


                    string temppath = System.IO.Path.GetTempPath();
                    string filename = "";
                    int fcount = 0;
                    string strfcount = "";
                    while(filename==""){
                        if (System.IO.File.Exists(temppath + id + strfcount + ".docx"))
                        {
                        try{
                            System.IO.File.Delete(temppath + id + strfcount + ".docx");
                            filename = temppath + id + strfcount + ".docx";
                        } catch(Exception){
                            fcount++;
                            strfcount = "_" + fcount.ToString() + "";
                        }
                    } else {
                        filename = temppath + id + strfcount + ".docx";
                    }
                    }
            return filename;
        }


        public static string SaveTempFile(string id,string filetype)
        {
            //Generate a temp file and save the doc there
            string temppath = System.IO.Path.GetTempPath();
            string filename = "";
            int fcount = 0;
            string strfcount = "";
            while (filename == "")
            {
                if (System.IO.File.Exists(temppath + id + strfcount + "." + filetype))
                {
                    try
                    {
                        System.IO.File.Delete(temppath + id + strfcount + "." + filetype);
                        filename = temppath + id + strfcount + "." + filetype;
                    }
                    catch (Exception)
                    {
                        fcount++;
                        strfcount = "_" + fcount.ToString() + "";
                    }
                }
                else
                {
                    filename = temppath + id + strfcount + "." + filetype;
                }
            }
            return filename;
        }

        public static string SaveTempHTMLFile(string id)
        {
            // Generate a temp file and save the doc there
            string temppath = System.IO.Path.GetTempPath();
            string filename = "";
            int fcount = 0;
            string strfcount = "";
            while (filename == "")
            {
                if (System.IO.File.Exists(temppath + id + strfcount + ".htm"))
                {
                    try
                    {
                        System.IO.File.Delete(temppath + id + strfcount + ".htm");
                        filename = temppath + id + strfcount + ".htm";
                    }
                    catch (Exception)
                    {
                        fcount++;
                        strfcount = "_" + fcount.ToString() + "";
                    }
                }
                else
                {
                    filename = temppath + id + strfcount + ".htm";
                }
            }
            return filename;
        }

        // Routine to handle excpetions when getting Data 
        // Put here so that we can hanlde the UI depending on the form we are in
        public static DataReturn HandleData(DataReturn dr)
        {
            if (dr == null) return null;
            if (!dr.success)
            {
                //if Message is TimeOut then warn and put up the login - otherwise show error
                if (dr.errormessage.Contains("INVALID_SESSION_ID"))
                {
                    MessageBox.Show("Timed Out! Please login again");
                    Globals.Ribbons.Ribbon1.Logout();
                    Globals.ThisAddIn.ProcessingStop("");
                }
                else
                {
                    Globals.ThisAddIn.ProcessingStop("");
                    MessageBox.Show("Sorry there has been an error:" + dr.errormessage);
                }
                Globals.ThisAddIn.ProcessingStop("");
            }
            return dr;
        }

        public static string ToText(long n)
        {
            return _toText(n, true);
        }
        private static string _toText(long n, bool isFirst = false)
        {
            string result;
            if (isFirst && n == 0)
            {
                result = "Zero";
            }
            else if (n < 0)
            {
                result = "Negative " + _toText(-n);
            }
            else if (n == 0)
            {
                result = "";
            }
            else if (n <= 9)
            {
                result = new[] { "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" }[n - 1] + " ";
            }
            else if (n <= 19)
            {
                result = new[] { "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" }[n - 10] + (isFirst ? null : " ");
            }
            else if (n <= 99)
            {
                result = new[] { "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" }[n / 10 - 2] + (n % 10 > 0 ? "-" + _toText(n % 10) : null);
            }
            else if (n <= 999)
            {
                result = _toText(n / 100) + "Hundred " + _toText(n % 100);
            }
            else if (n <= 999999)
            {
                result = _toText(n / 1000) + "Thousand " + _toText(n % 1000);
            }
            else if (n <= 999999999)
            {
                result = _toText(n / 1000000) + "Million " + _toText(n % 1000000);
            }
            else
            {
                result = _toText(n / 1000000000) + "Billion " + _toText(n % 1000000000);
            }
            if (isFirst)
            {
                result = result.Trim();
            }
            return result;
        }


        public static void UnlockContentControls(Word.Document doc)
        {
            Word.Range r = doc.Range(doc.Content.Start, doc.Content.End);

            Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                cc.LockContents = false;
                cc.LockContentControl = false;
            }
        }

        public static void RemoveContentControls(Word.Document doc)
        {
            Word.Range r = doc.Range(doc.Content.Start, doc.Content.End);

            Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                cc.LockContents = false;
                cc.LockContentControl = false;

                Word.Range temp = cc.Range;
                string text = temp.Text;
                if (text == "") text = Convert.ToString(cc.Title);
                cc.Delete();
                temp.Text = "";
                temp.InsertBefore(text);
            }
        }

        public static Word.Range RemoveElements(Word.Range r)
        {          
            Word.ContentControl[] ccs = new Word.ContentControl[r.ContentControls.Count];
            foreach (Word.ContentControl cc in r.ContentControls)
            {
                cc.LockContents = false;
                cc.LockContentControl = false;
                Word.Range temp = cc.Range;
                cc.Delete();
                temp.Text = "";
            }
            return r;
        }


        public static void setTheme(DependencyObject node)
        {
            StyleManager.SetTheme(node, StyleManager.ApplicationTheme);
            List<DependencyObject> l = GetLogicalChildCollection<DependencyObject>(node);
            foreach (DependencyObject o in l)
            {
                StyleManager.SetTheme(o, StyleManager.ApplicationTheme);
            }
        }

        public static List<T> GetLogicalChildCollection<T>(object parent) where T : DependencyObject
        {
            List<T> logicalCollection = new List<T>();
            GetLogicalChildCollection(parent as DependencyObject, logicalCollection);
            return logicalCollection;
        }

        private static void GetLogicalChildCollection<T>(DependencyObject parent, List<T> logicalCollection) where T : DependencyObject
        {
            IEnumerable children = LogicalTreeHelper.GetChildren(parent);
            foreach (object child in children)
            {
                if (child is DependencyObject)
                {
                    DependencyObject depChild = child as DependencyObject;
                    if (child is T)
                    {
                        logicalCollection.Add(child as T);
                    }
                    GetLogicalChildCollection(depChild, logicalCollection);
                }
            }
        }


        //---- Encrypt/Decrypt functions lifted from Stackoverflow - not meant to be hugely secure 
        //---- just used to secure demo passwords

         // This constant string is used as a "salt" value for the PasswordDeriveBytes function calls.
        // This size of the IV (in bytes) must = (keysize / 8).  Default keysize is 256, so the IV must be
        // 32 bytes long.  Using a 16 character string here gives us 32 bytes when converted to a byte array.
        private static readonly byte[] initVectorBytes = Encoding.ASCII.GetBytes("qx2f1jxx1faqzgd5");

        // This constant is used to determine the keysize of the encryption algorithm.
        private const int keysize = 256;

        public static string Encrypt(string plainText, string passPhrase)
        {
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            using (PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null))
            {
                byte[] keyBytes = password.GetBytes(keysize / 8);
                using (RijndaelManaged symmetricKey = new RijndaelManaged())
                {
                    symmetricKey.Mode = CipherMode.CBC;
                    using (ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes))
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            using (CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                            {
                                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                                cryptoStream.FlushFinalBlock();
                                byte[] cipherTextBytes = memoryStream.ToArray();
                                return Convert.ToBase64String(cipherTextBytes);
                            }
                        }
                    }
                }
            }
        }

        public static string Decrypt(string cipherText, string passPhrase)
        {
            try
            {
                byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
                using (PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null))
                {
                    byte[] keyBytes = password.GetBytes(keysize / 8);
                    using (RijndaelManaged symmetricKey = new RijndaelManaged())
                    {
                        symmetricKey.Mode = CipherMode.CBC;
                        using (ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes))
                        {
                            using (MemoryStream memoryStream = new MemoryStream(cipherTextBytes))
                            {
                                using (CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                                {
                                    byte[] plainTextBytes = new byte[cipherTextBytes.Length];
                                    int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
                                    return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
                                }
                            }
                        }
                    }
                }
            } catch(Exception){
                //If the string isn't in base 64 jsut pass it back in plain text
                return cipherText;
            }
        }


        public static string Truncate(string x, int maxLength)
        {
            if (string.IsNullOrEmpty(x))
            {
                return x;
            }
            else if (x.Length <= maxLength)
            {
                return x;
            }
            else
            {
                return x.Substring(0, maxLength);
            }
        }


        public static string FixUpSOQLString(string x)
        {
            if (x == null) return x;
            return x.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\"", "\\\"");
        }

        public static string CleanUpXML(string val)
        {
            string rtn = Regex.Replace(val, @"[\x00-\x08]|[\x0B\x0C]|[\x0E-\x19]|[\uD800-\uDFFF]|[\uFFFE\uFFFF]", "");
            rtn = Regex.Replace(rtn, @"[\x1A-\x1F]", "-");
            return rtn;
        }
    }
    
}
