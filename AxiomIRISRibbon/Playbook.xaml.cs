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
using System.IO;
using System.Xml;
using System.Windows.Markup;
using System.Diagnostics;
using HTMLConverter;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for Playbook.xaml
    /// </summary>
    public partial class Playbook : Window
    {
        private Data D;

        private string Id;
        private string Type;

        private string InfoHtml;
        private string ClientHtml;

        private TemplateEdit.TEditSidebar TemplateSideBar;
        private bool HasChanged;

        public Playbook()
        {
            InitializeComponent();
            Utility.setTheme(this);

            AddHandler(Hyperlink.RequestNavigateEvent, new RoutedEventHandler(OnNavigationRequest));

            this.D = Globals.ThisAddIn.getData();

            this.btnOK.IsEnabled = false;
        }

        public void OnNavigationRequest(object sender, RoutedEventArgs e)
        {
            var source = e.OriginalSource as Hyperlink;
            if (source != null)
                Process.Start(source.NavigateUri.ToString());
        }

        // Open from the Template Sidebar
        public void Open(TemplateEdit.TEditSidebar ts,string ConceptId, string Html, string pbType)
        {
            this.TemplateSideBar = ts;
            this.HasChanged = false;
            this.Open(ConceptId, Html, pbType);

            this.btnEdit.Visibility = System.Windows.Visibility.Visible;
            this.btnFootnotes.Visibility = System.Windows.Visibility.Visible;
        }

        // Open from the Contract Sidebar
        public void OpenFromContract(string ConceptId, string Html, string pbType)
        {
            this.TemplateSideBar = null;
            this.HasChanged = false;
            this.Open(ConceptId, Html, pbType);

            this.btnEdit.Visibility = System.Windows.Visibility.Hidden;
            this.btnFootnotes.Visibility = System.Windows.Visibility.Hidden;
        }


        //Open when passed the html
        public void Open(string ConceptId, string Html,string pbType)
        {
            this.Id = ConceptId;
            this.Type = pbType;


            if (Html != "")
            {

                StringReader stringReader = new StringReader(HtmlToXamlConverter.ConvertHtmlToXaml(Html, true));
                XmlReader xmlReader = XmlReader.Create(stringReader);
                FlowDocument fdoc = (FlowDocument)XamlReader.Load(xmlReader);
                this.richTextBox1.Document = fdoc;
            }
            else
            {
                this.richTextBox1.SelectAll();
                this.richTextBox1.Selection.Text = "";
            }

            this.btnOK.IsEnabled = false;

        }


        //Open if we don't have the HTML
        public void Open(string ConceptId, string pbType)
        {
            this.Id = ConceptId;
            this.Type = pbType;

            DataTable c = Utility.HandleData(this.D.GetConcept(ConceptId)).dt;
            if (c.Rows.Count > 0)
            {
                string html = "";
                this.InfoHtml = c.Rows[0]["PlayBookInfo__c"].ToString();
                this.ClientHtml = c.Rows[0]["PlayBookClient__c"].ToString();

                if (this.Type == "Info")
                {
                    html = this.InfoHtml;
                }
                else
                {
                    html = this.ClientHtml;
                }

                if (html != "")
                {

                    StringReader stringReader = new StringReader(HtmlToXamlConverter.ConvertHtmlToXaml(html, true));
                    XmlReader xmlReader = XmlReader.Create(stringReader);
                    FlowDocument fdoc = (FlowDocument)XamlReader.Load(xmlReader);
                    this.richTextBox1.Document = fdoc;
                }
                else
                {
                    this.richTextBox1.SelectAll();
                    this.richTextBox1.Selection.Text = "";
                }

            }

            this.btnOK.IsEnabled = false;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (this.HasChanged)
            {
                TextRange range = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd);
                MemoryStream stream = new MemoryStream();
                range.Save(stream, DataFormats.Xaml);
                string xamlText = Encoding.UTF8.GetString(stream.ToArray());
                if (!xamlText.Trim().ToLower().StartsWith("<flowdocument>")) xamlText = "<FlowDocument>" + xamlText + "</FlowDocument>";
                string html = HtmlFromXamlConverter.ConvertXamlToHtml(xamlText);
                if (this.Type == "Info")
                {
                    this.InfoHtml = html;
                    this.ClientHtml = null;
                }
                else if (this.Type == "Client")
                {
                    this.InfoHtml = null;
                    this.ClientHtml = html;
                }

                this.D.SaveConcept(this.Id, this.InfoHtml, this.ClientHtml);

                // if we have been called from the Template Sidebar then get it to update
                if (this.TemplateSideBar != null)
                {
                    this.TemplateSideBar.UpdateCachePlaybook(this.Id,this.Type, html);
                }
            }


            this.Hide();

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (this.richTextBox1.IsReadOnly)
            {
                this.richTextBox1.IsReadOnly = false;
                this.imgLock.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/unlocksmall.png", UriKind.Relative));
            }
            else
            {
                this.richTextBox1.IsReadOnly = true;
                this.imgLock.Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/AxiomIRISRibbon;component/Resources/locksmall.png", UriKind.Relative));
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }

        private void richTextBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.HasChanged = true;
            this.btnOK.IsEnabled = true;
        }

        private void btnFootnotes_Click(object sender, RoutedEventArgs e)
        {
            string footnotes = this.TemplateSideBar.GetFootnotes();
            if (footnotes != "")
            {
                this.richTextBox1.SelectAll();
                if (this.richTextBox1.Selection.Text.Trim() != "") this.richTextBox1.AppendText("\n");
                this.richTextBox1.AppendText(footnotes);
            }
        }
    }
}
