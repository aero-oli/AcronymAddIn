using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace AcronymAddIn
{
    public partial class AcronymSelectionForm : Form
    {
        private List<string> _acronyms;
        private Dictionary<string, List<string>> _mappings;
        private FlowLayoutPanel panel;
        private Button btnGenerate;
        private TextBox txtCsvPath;
        private Button btnBrowse;

        public AcronymSelectionForm(List<string> acronyms, Dictionary<string, List<string>> mappings)
        {
            _acronyms = acronyms;
            _mappings = mappings;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Select Acronym Meanings";
            this.Size = new System.Drawing.Size(400, 600);

            panel = new FlowLayoutPanel()
            {
                Dock = DockStyle.Top,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Height = 500
            };
            this.Controls.Add(panel);

            // Create a UI group for each detected acronym.
            foreach (var acronym in _acronyms)
            {
                GroupBox groupBox = new GroupBox() { Text = acronym, Width = 350, Height = 60 };
                ComboBox comboBox = new ComboBox() { Left = 10, Top = 20, Width = 200 };
                if (_mappings.ContainsKey(acronym))
                    comboBox.Items.AddRange(_mappings[acronym].ToArray());
                CheckBox checkBox = new CheckBox() { Left = 220, Top = 22, Text = "Include" };
                groupBox.Controls.Add(comboBox);
                groupBox.Controls.Add(checkBox);
                panel.Controls.Add(groupBox);
            }

            txtCsvPath = new TextBox() { Left = 10, Top = 520, Width = 300 };
            btnBrowse = new Button() { Text = "Browse", Left = 320, Top = 518, Width = 60 };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(txtCsvPath);
            this.Controls.Add(btnBrowse);

            btnGenerate = new Button() { Text = "Generate Table", Dock = DockStyle.Bottom };
            btnGenerate.Click += BtnGenerate_Click;
            this.Controls.Add(btnGenerate);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtCsvPath.Text = openFileDialog.FileName;
                }
            }
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            var selectedData = new List<(string Acronym, string Meaning)>();
            foreach (GroupBox group in panel.Controls)
            {
                string acronym = group.Text;
                var combo = group.Controls.OfType<ComboBox>().FirstOrDefault();
                var check = group.Controls.OfType<CheckBox>().FirstOrDefault();
                if (check != null && check.Checked && combo != null && combo.SelectedItem != null)
                {
                    selectedData.Add((acronym, combo.SelectedItem.ToString()));
                }
            }
            InsertAcronymTable(selectedData);
            this.Close();
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
    }
}