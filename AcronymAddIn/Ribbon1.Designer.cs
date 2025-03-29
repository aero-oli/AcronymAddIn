namespace AcronymAddIn
{
    partial class Ribbon1
    {
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnDetectAcronyms;
        private Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateAcronyms;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDetectAcronyms = this.Factory.CreateRibbonButton();
            this.btnUpdateAcronyms = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Acronym Tools";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDetectAcronyms);
            this.group1.Items.Add(this.btnUpdateAcronyms);
            this.group1.Label = "Acronym Operations";
            this.group1.Name = "group1";
            // 
            // btnDetectAcronyms
            // 
            this.btnDetectAcronyms.Label = "Detect Acronyms";
            this.btnDetectAcronyms.Name = "btnDetectAcronyms";
            this.btnDetectAcronyms.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDetectAcronyms_Click);
            // 
            // btnUpdateAcronyms
            // 
            this.btnUpdateAcronyms.Label = "Update Acronyms";
            this.btnUpdateAcronyms.Name = "btnUpdateAcronyms";
            this.btnUpdateAcronyms.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateAcronyms_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
        }

        #endregion
    }
}