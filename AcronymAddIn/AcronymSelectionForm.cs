using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace AcronymAddIn
{
    public partial class AcronymSelectionForm : Form
    {
        private List<string> _acronyms;
        private Dictionary<string, List<string>> _mappings;
        private FlowLayoutPanel panel;
        private Button btnGenerate;
        private Button btnOpenCsv;
        private Word.Table _existingTable;

        public AcronymSelectionForm(List<string> acronyms, Dictionary<string, List<string>> mappings, Word.Table existingTable = null)
        {
            _acronyms = acronyms;
            _mappings = mappings;
            _existingTable = existingTable;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Select Acronym Meanings";
            this.Size = new System.Drawing.Size(500, 700);
            this.MinimumSize = new System.Drawing.Size(500, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            panel = new FlowLayoutPanel()
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10),
                Margin = new Padding(10)
            };
            this.Controls.Add(panel);

            // Create a UI group for each detected acronym.
            foreach (var acronym in _acronyms)
            {
                GroupBox groupBox = new GroupBox() { Text = acronym, Width = 450, Height = 80, Padding = new Padding(10), Margin = new Padding(5) };
                ComboBox comboBox = new ComboBox() 
                { 
                    Left = 10, 
                    Top = 20, 
                    Width = 300, 
                    DropDownStyle = ComboBoxStyle.DropDown, 
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                    DropDownWidth = 450,
                    AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                    AutoCompleteSource = AutoCompleteSource.ListItems
                };
                if (_mappings.ContainsKey(acronym))
                    comboBox.Items.AddRange(_mappings[acronym].ToArray());
                CheckBox checkBox = new CheckBox() 
                { 
                    Text = "Include", 
                    Left = comboBox.Right + 10, 
                    Top = 22, 
                    Anchor = AnchorStyles.Top | AnchorStyles.Right 
                };
                groupBox.Controls.Add(comboBox);
                groupBox.Controls.Add(checkBox);
                panel.Controls.Add(groupBox);
            }

            btnOpenCsv = new Button() { Text = "Open CSV", Dock = DockStyle.Bottom, Height = 40, Margin = new Padding(10) };
            btnOpenCsv.Click += BtnOpenCsv_Click;
            this.Controls.Add(btnOpenCsv);

            btnGenerate = new Button() 
            { 
                Text = _existingTable != null ? "Update Table" : "Generate Table", 
                Dock = DockStyle.Bottom, 
                Height = 40, 
                Margin = new Padding(10) 
            };
            btnGenerate.Click += BtnGenerate_Click;
            this.Controls.Add(btnGenerate);
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            var selectedData = new List<(string Acronym, string Meaning)>();
            foreach (GroupBox group in panel.Controls)
            {
                string acronym = group.Text;
                var combo = group.Controls.OfType<ComboBox>().FirstOrDefault();
                var check = group.Controls.OfType<CheckBox>().FirstOrDefault();
                if (check != null && check.Checked && combo != null && combo.Text != null)
                {
                    string selectedMeaning = combo.Text;
                    selectedData.Add((acronym, selectedMeaning));
                    SaveAcronymToCsv(acronym, selectedMeaning);
                }
            }

            if (_existingTable != null)
            {
                // Update existing table
                foreach (var (Acronym, Meaning) in selectedData)
                {
                    var newRow = _existingTable.Rows.Add();
                    newRow.Cells[1].Range.Text = Acronym;
                    newRow.Cells[2].Range.Text = Meaning;
                }
            }
            else
            {
                // Create new table
                InsertAcronymTable(selectedData);
            }
            this.Close();
        }

        private void BtnOpenCsv_Click(object sender, EventArgs e)
        {
            string homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string csvFilePath = Path.Combine(homeDirectory, "acronyms.csv");
            if (File.Exists(csvFilePath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = csvFilePath,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("CSV file not found.");
            }
        }

        private void SaveAcronymToCsv(string acronym, string meaning)
        {
            string homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string csvFilePath = Path.Combine(homeDirectory, "acronyms.csv");

            // Read existing entries
            var existingEntries = new HashSet<string>();
            if (File.Exists(csvFilePath))
            {
                var lines = File.ReadLines(csvFilePath);
                foreach (var line in lines)
                {
                    existingEntries.Add(line.Trim());
                }
            }
            string newEntry = $"{acronym},{meaning}";

            // Check if the entry already exists
            if (!existingEntries.Contains(newEntry))
            {
                using (StreamWriter writer = new StreamWriter(csvFilePath, true))
                {
                    writer.WriteLine(newEntry);
                }
            }
        }

        // Inserts a table into the active Word document at the current selection.
        private void InsertAcronymTable(List<(string Acronym, string Meaning)> data)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            var range = Globals.ThisAddIn.Application.Selection.Range;
            int rows = data.Count + 1;
            int cols = 2;
            var table = doc.Tables.Add(range, rows, cols);
            table.Cell(1, 1).Range.Text = "Acronym";
            table.Cell(1, 2).Range.Text = "Meaning";
            for (int i = 0; i < data.Count; i++)
            {
                table.Cell(i + 2, 1).Range.Text = data[i].Acronym;
                table.Cell(i + 2, 2).Range.Text = data[i].Meaning;
            }
        }

        private void UpdateAcronymTable()
        {
            // This method is no longer needed as the update functionality is handled differently
        }

        private List<string> DetectAcronyms(string text)
        {
            var matches = System.Text.RegularExpressions.Regex.Matches(text, @"\b[A-Z]{3,}\b");
            return matches.Cast<System.Text.RegularExpressions.Match>().Select(m => m.Value).Distinct().ToList();
        }
    }
}