using Microsoft.Win32;
using System;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.CodeDom;
using System.Xml;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Http;
using System.Collections;

namespace WordLinkValidator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private string[] filesToAnalyze;

        public MainWindow()
        {
            filesToAnalyze = Array.Empty<string>();

            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog {
                Multiselect = true,
                Filter      = "Word files|*.doc;*.docx"
            };
			
            if (openFileDialog.ShowDialog() != true) {
                return;
            }

            filesToAnalyze = openFileDialog.FileNames;
        }

        private void btnValidate_Click(object sender, RoutedEventArgs e)
        {
            if (filesToAnalyze == null || filesToAnalyze.Length < 1) {
                MessageBox.Show("Please select Word files to analyze.", "Missing Documents", MessageBoxButton.OK);

                return;
            }

            dataGridWordList.Items.Clear();
            dataGridWordList.Items.Refresh();

            foreach (string file in filesToAnalyze) {
                WordprocessingDocument wordDocument = WordprocessingDocument.Open(file, false);

                if (wordDocument.MainDocumentPart == null) {
                    continue;
                }

                System.IO.StreamReader sr = new System.IO.StreamReader(
                    wordDocument.MainDocumentPart.GetStream()
                );
                string xml = sr.ReadToEnd();

                // Handle content links
                foreach (HyperlinkRelationship link 
                    in wordDocument.MainDocumentPart.HyperlinkRelationships
                ) {
                    this.handleLink("Content", link, file, xml);
                }

                // Handle header/footer links
                string headerFooterRelationshipId;
                foreach (HeaderFooterReferenceType headerFooterRef 
                    in wordDocument.MainDocumentPart.Document.Descendants<HeaderFooterReferenceType>()
                ) {
                    if (headerFooterRef == null || headerFooterRef.Id == null || headerFooterRef.Id.Value == null) {
                        continue;
                    }

                    headerFooterRelationshipId = headerFooterRef.Id.Value;
                    foreach (HyperlinkRelationship hfLink 
                        in wordDocument.MainDocumentPart.GetPartById(headerFooterRelationshipId).HyperlinkRelationships
                    ) {
                        this.handleLink("Header/Footer", hfLink, file, xml);
                    }
                }

                if (wordDocument.MainDocumentPart.WordprocessingCommentsPart != null) {
                    foreach (HyperlinkRelationship link
                        in wordDocument.MainDocumentPart.WordprocessingCommentsPart.HyperlinkRelationships
                    ) {
                        this.handleLink("Comment", link, file, xml);
                    }
                }

                if (wordDocument.MainDocumentPart.FootnotesPart != null) {
                    foreach (HyperlinkRelationship link
                        in wordDocument.MainDocumentPart.FootnotesPart.HyperlinkRelationships
                    ) {
                        this.handleLink("Footnote", link, file, xml);
                    }
                }

                if (wordDocument.MainDocumentPart.EndnotesPart != null) {
                    foreach (HyperlinkRelationship link
                        in wordDocument.MainDocumentPart.EndnotesPart.HyperlinkRelationships
                    ) {
                        this.handleLink("Endnote", link, file, xml);
                    }
                }
            }

            filesToAnalyze = Array.Empty<string>();
        }

        private async void handleLink(string location, HyperlinkRelationship link, string file, string xml)
        {
            string linkDestination = link.Uri.ToString();

            Regex rx = new Regex("(r:id=\\\"" + link.Id + "\\\".*?<w:t>)(.*?)(</w:t>)");
            MatchCollection matches = rx.Matches(xml);

            string linkName = "";
            if (matches.Count > 0 && matches[0].Groups.Count > 2) {
                linkName = matches[0].Groups[2].Value;
            }

            string status = linkIsReachable(linkDestination) ? "OK" : "ERROR";

            dataGridWordList.Items.Add(new
            {
                Status = status,
                Type = linkDestination.StartsWith("http") ? "URL" : "Local",
                Location = location,
                File = file,
                Name = linkName,
                Link = linkDestination
            });
        }

        private static bool linkIsReachable(string url)
        {
            HttpClient client = new HttpClient();
            client.Timeout    = TimeSpan.FromSeconds(5);

            try {
                HttpResponseMessage response = client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).Result;
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    return true;
                }
            } catch (Exception) {
                return false;
            }

            return false;
        }
    }
}
