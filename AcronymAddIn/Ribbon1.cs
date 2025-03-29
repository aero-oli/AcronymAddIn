using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace AcronymAddIn
{
    public partial class Ribbon1 : RibbonBase
    {
        // Single constructor for Ribbon1.
        public Ribbon1() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Additional load code, if necessary.
        }

        // Event handler for the custom ribbon button.
        private void btnDetectAcronyms_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string docText = doc.Content.Text;

            // Detect acronyms (three or more consecutive uppercase letters).
            List<string> acronyms = DetectAcronyms(docText);

            // Load CSV mappings from the path specified in App.config.
            string csvPath = ConfigurationManager.AppSettings["AcronymCsvPath"];
            Dictionary<string, List<string>> acronymMappings = LoadAcronymMappings(csvPath);

            // Show the dynamic form for acronym selection.
            AcronymSelectionForm form = new AcronymSelectionForm(acronyms, acronymMappings);
            form.ShowDialog();
        }

        // Uses a regex to detect acronyms in the text.
        private List<string> DetectAcronyms(string text)
        {
            var matches = Regex.Matches(text, @"\b[A-Z]{3,}\b");
            return matches.Cast<Match>().Select(m => m.Value).Distinct().ToList();
        }

        // Loads acronym-to-meanings mappings from a CSV file.
        private Dictionary<string, List<string>> LoadAcronymMappings(string csvFilePath)
        {
            var mappings = new Dictionary<string, List<string>>();
            if (File.Exists(csvFilePath))
            {
                string[] lines = File.ReadAllLines(csvFilePath);
                foreach (var line in lines)
                {
                    var parts = line.Split(',');
                    if (parts.Length >= 2)
                    {
                        string acronym = parts[0].Trim();
                        string meaning = parts[1].Trim();
                        if (!mappings.ContainsKey(acronym))
                            mappings[acronym] = new List<string>();
                        mappings[acronym].Add(meaning);
                    }
                }
            }
            else
            {
                MessageBox.Show("CSV file not found: " + csvFilePath);
            }
            return mappings;
        }
    }
}