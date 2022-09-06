using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Windows;

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
                Filter      = "Word files (*.doc)|*.doc|Word files (*.docx)|*.docx"
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

            string text                                    = "";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            foreach (string file in filesToAnalyze) {
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(file);

                // Paragraphs
                foreach (Microsoft.Office.Interop.Word.Paragraph para in doc.Paragraphs) {
                    Microsoft.Office.Interop.Word.Range range = para.Range;
                    text                                      = range.Text;

                    if (hasLink(text)) {
                        string[][] links = findLinks(text);

                        foreach (string[] link in links) {
                            // @TODO: validate link
                            // @TODO: add validation result to list
                        }
                    }
                }

                // Footnotes
                foreach (Microsoft.Office.Interop.Word.Footnote footnote in doc.Footnotes) {
                }

                // Endings
                foreach (Microsoft.Office.Interop.Word.Endnote ending in doc.Endnotes) {
                }

                // doc.Hyperlinks
               

                // @TODO: check header and footer

                doc.Close();
            }

            word.Quit();

            filesToAnalyze = Array.Empty<string>();
        }

        private static bool hasLink(string text)
        {
            return text.Contains("http://")
                || text.Contains("https://")
                || text.Contains("file://");
        }

        private static string[][] findLinks(string text)
        {
            return new string[][] { Array.Empty<string>() };
        }
    }
}
