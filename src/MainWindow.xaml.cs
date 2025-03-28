﻿using Microsoft.Win32;
using System;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using WordLinkValidator.Views;

namespace WordLinkValidator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private string[] filesToAnalyze;
        private int linkCount        = 0;
        private int invalidLinkCount = 0;
        private int fileCount        = 0;

        public MainWindow()
        {
            filesToAnalyze = Array.Empty<string>();

            InitializeComponent();

            #if OMS_DEMO
                this.Title = "Demo " + this.Title;
                MessageBox.Show("This is a demo with limited functionality.");
            #endif
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

            linkCount        = 0;
            invalidLinkCount = 0;
            fileCount        = filesToAnalyze.Length;

            dataGridWordList.Items.Clear();
            dataGridWordList.Items.Refresh();

            // Iterate all selected files
            foreach (string file in filesToAnalyze) {
                WordprocessingDocument wordDocument = WordprocessingDocument.Open(file, false);

                if (wordDocument.MainDocumentPart == null) {
                    continue;
                }

                // Handle content links
                StreamReader sr = new StreamReader(
                    wordDocument.MainDocumentPart.GetStream()
                );
                string xml = sr.ReadToEnd();

                foreach (HyperlinkRelationship link 
                    in wordDocument.MainDocumentPart.HyperlinkRelationships
                ) {
                    this.handleLink("Content", link, file, xml);
                }

                // Handle header/footer links
                string headerFooterRelationshipId;
                OpenXmlPart headerFooterPart;
                foreach (HeaderFooterReferenceType headerFooterRef 
                    in wordDocument.MainDocumentPart.Document.Descendants<HeaderFooterReferenceType>()
                ) {
                    if (headerFooterRef == null || headerFooterRef.Id == null || headerFooterRef.Id.Value == null) {
                        continue;
                    }

                    headerFooterRelationshipId = headerFooterRef.Id.Value;
                    headerFooterPart           = wordDocument.MainDocumentPart.GetPartById(headerFooterRelationshipId);

                    sr = new StreamReader(
                        headerFooterPart.GetStream()
                    );
                    xml = sr.ReadToEnd();
                    
                    foreach (HyperlinkRelationship hfLink in headerFooterPart.HyperlinkRelationships) {
                        this.handleLink("Header/Footer", hfLink, file, xml);
                    }
                }

                // Handle comment links
                if (wordDocument.MainDocumentPart.WordprocessingCommentsPart != null) {
                    sr = new StreamReader(
                        wordDocument.MainDocumentPart.WordprocessingCommentsPart.GetStream()
                    );
                    xml = sr.ReadToEnd();

                    foreach (HyperlinkRelationship link
                        in wordDocument.MainDocumentPart.WordprocessingCommentsPart.HyperlinkRelationships
                    ) {
                        this.handleLink("Comment", link, file, xml);
                    }
                }

                // Handle footnote links
                if (wordDocument.MainDocumentPart.FootnotesPart != null) {
                    sr = new StreamReader(
                        wordDocument.MainDocumentPart.FootnotesPart.GetStream()
                    );
                    xml = sr.ReadToEnd();

                    foreach (HyperlinkRelationship link
                        in wordDocument.MainDocumentPart.FootnotesPart.HyperlinkRelationships
                    ) {
                        this.handleLink("Footnote", link, file, xml);
                    }
                }

                // Handle endnote links
                if (wordDocument.MainDocumentPart.EndnotesPart != null) {
                    sr = new StreamReader(
                        wordDocument.MainDocumentPart.EndnotesPart.GetStream()
                    );
                    xml = sr.ReadToEnd();

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

            // Get link name
            Regex rx = new Regex("(r:id=\\\"" + link.Id + "\\\".*?<w:t>)(.*?)(</w:t>)");
            MatchCollection matches = rx.Matches(xml);

            string linkName = "";
            if (matches.Count > 0 && matches[0].Groups.Count > 2) {
                linkName = matches[0].Groups[2].Value;
            }

            // Check link status
            string status = (await linkIsReachable(linkDestination, Path.GetDirectoryName(file))) ? "OK" : "ERROR";

            #if OMS_DEMO
                if (linkCount > 1) {
                    return;
                }
            #endif

            // Change stats
            ++linkCount;
            if (status == "ERROR") {
                ++invalidLinkCount;
            }

            // Add item to grid
            dataGridWordList.Items.Add(new
            {
                Status = status,
                Type = linkDestination.StartsWith("http") ? "URL" : "Local",
                Location = location,
                File = Path.GetFileName(file),
                Name = linkName,
                Link = linkDestination
            });

            textBlockStatus.Text = "Files: " + fileCount
                + " Links: " + linkCount
                + " Invalid: " + invalidLinkCount;
        }

        private static async Task<bool> linkIsReachable(string url, string? basePath)
        {
            if (url.StartsWith("http") || url.StartsWith("www")) {
                // Handle web link
                HttpClient client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(5);

                return await Task.Run(() =>
                {
                    try
                    {
                        HttpResponseMessage response = client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).Result;
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                        return false;
                    }

                    return false;
                });
            } else {
                // Handle local file link
                return await Task.Run(() =>
                {
                    return File.Exists(url) || Directory.Exists(url)
                        || File.Exists(basePath + "/" + url) || Directory.Exists(basePath + "/" + url);
                });
            }
        }

        private void menuExit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
            Environment.Exit(0);
        }

        private void menuInfo_Click(object sender, RoutedEventArgs e)
        {
            if (Info.isOpen) {
                return;
            }

            Info window = new Info();
            window.Show();
        }
    }
}
