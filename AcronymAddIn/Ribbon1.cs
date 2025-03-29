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

        private void EnsureCsvFileExists()
        {
            string homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string csvFilePath = Path.Combine(homeDirectory, "acronyms.csv");

            if (!File.Exists(csvFilePath))
            {
                using (StreamWriter writer = new StreamWriter(csvFilePath))
                {
                    writer.WriteLine("Acronym,Meaning");
                    writer.WriteLine("NASA,National Aeronautics and Space Administration");
                }
            }
        }

        // Event handler for the custom ribbon button.
        private void btnDetectAcronyms_Click(object sender, RibbonControlEventArgs e)
        {
            EnsureCsvFileExists();
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string docText = doc.Content.Text;

            // Detect acronyms (three or more consecutive uppercase letters).
            List<string> acronyms = DetectAcronyms(docText);

            // Load CSV mappings from the home directory.
            string homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string csvPath = Path.Combine(homeDirectory, "acronyms.csv");
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

        // Event handler for updating existing acronym tables
        private void btnUpdateAcronyms_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                
                // Try to find an acronym table in the document
                Word.Table acronymTable = null;
                foreach (Word.Table table in doc.Tables)
                {
                    if (table.Rows.Count > 0 && 
                        table.Columns.Count == 2)
                    {
                        string header1 = table.Cell(1, 1).Range.Text.Replace("\r", "").Replace("\a", "").Trim();
                        string header2 = table.Cell(1, 2).Range.Text.Replace("\r", "").Replace("\a", "").Trim();
                        
                        if (header1 == "Acronym" && header2 == "Meaning")
                        {
                            acronymTable = table;
                            break;
                        }
                    }
                }

                if (acronymTable == null)
                {
                    MessageBox.Show("No acronym table found in the document. Please create a table first using 'Detect Acronyms'.");
                    return;
                }

                // Get existing acronyms from the table
                var existingAcronyms = new Dictionary<string, string>();
                for (int i = 2; i <= acronymTable.Rows.Count; i++)
                {
                    string acronym = acronymTable.Cell(i, 1).Range.Text.Replace("\r", "").Replace("\a", "").Trim();
                    string meaning = acronymTable.Cell(i, 2).Range.Text.Replace("\r", "").Replace("\a", "").Trim();
                    if (!string.IsNullOrWhiteSpace(acronym))
                    {
                        existingAcronyms[acronym] = meaning;
                    }
                }

                // Detect all acronyms in the document
                string docText = doc.Content.Text;
                List<string> allAcronyms = DetectAcronyms(docText);
                
                // Filter out acronyms that are already in the table
                List<string> newAcronyms = allAcronyms.Where(a => !existingAcronyms.ContainsKey(a)).ToList();

                if (newAcronyms.Count == 0)
                {
                    MessageBox.Show($"No new acronyms found in the document.\nCurrent acronyms in table: {string.Join(", ", existingAcronyms.Keys)}");
                    return;
                }

                // Load mappings and show form for new acronyms only
                EnsureCsvFileExists();
                string homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string csvPath = Path.Combine(homeDirectory, "acronyms.csv");
                Dictionary<string, List<string>> acronymMappings = LoadAcronymMappings(csvPath);

                string currentAcronyms = string.Join(", ", existingAcronyms.Keys);
                string newAcronymsStr = string.Join(", ", newAcronyms);
                
                var result = MessageBox.Show(
                    $"Current acronyms in table: {currentAcronyms}\n\n" +
                    $"New acronyms found: {newAcronymsStr}\n\n" +
                    "Would you like to add these new acronyms to the table?",
                    "Update Acronyms",
                    MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    var form = new AcronymSelectionForm(newAcronyms, acronymMappings, acronymTable);
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating acronyms: {ex.Message}");
            }
        }
    }
}