namespace SQLServerForExcel_Addin
{
    partial class ExcelAddinTaskPane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelAddinTaskPane));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnConnectToDatabase = new System.Windows.Forms.ToolStripButton();
            this.btnRefreshData = new System.Windows.Forms.ToolStripButton();
            this.btnApplyChangesToDb = new System.Windows.Forms.ToolStripButton();
            this.btnSaveChangesToFile = new System.Windows.Forms.ToolStripButton();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabDatabaseExplorer = new System.Windows.Forms.TabPage();
            this.tvTables = new System.Windows.Forms.TreeView();
            this.tabPageSheetChanges = new System.Windows.Forms.TabPage();
            this.lblRefresh = new System.Windows.Forms.LinkLabel();
            this.lvSheetChanges = new System.Windows.Forms.ListView();
            this.chPrimaryKey = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chColName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chNewValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tabPageDataGeneration = new System.Windows.Forms.TabPage();
            this.btnInsertDataToSelection = new System.Windows.Forms.Button();
            this.cboColumnNames = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnBrowseForDataFile = new System.Windows.Forms.Button();
            this.txtDataFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.diagOpenFile = new System.Windows.Forms.OpenFileDialog();
            this.diagSaveFile = new System.Windows.Forms.SaveFileDialog();
            this.toolStrip1.SuspendLayout();
            this.tabMain.SuspendLayout();
            this.tabDatabaseExplorer.SuspendLayout();
            this.tabPageSheetChanges.SuspendLayout();
            this.tabPageDataGeneration.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnConnectToDatabase,
            this.btnRefreshData,
            this.btnApplyChangesToDb,
            this.btnSaveChangesToFile});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(322, 27);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // btnConnectToDatabase
            // 
            this.btnConnectToDatabase.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnConnectToDatabase.Image = global::ExcelAddIn1.Properties.Resources.MindMapImportData;
            this.btnConnectToDatabase.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnConnectToDatabase.Name = "btnConnectToDatabase";
            this.btnConnectToDatabase.Size = new System.Drawing.Size(29, 24);
            this.btnConnectToDatabase.Text = "Connect to Database";
            this.btnConnectToDatabase.Click += new System.EventHandler(this.btnConnectToDatabase_Click);
            // 
            // btnRefreshData
            // 
            this.btnRefreshData.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnRefreshData.Image = global::ExcelAddIn1.Properties.Resources.RefreshData;
            this.btnRefreshData.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnRefreshData.Name = "btnRefreshData";
            this.btnRefreshData.Size = new System.Drawing.Size(29, 24);
            this.btnRefreshData.Text = "Refresh Data";
            this.btnRefreshData.Click += new System.EventHandler(this.btnRefreshData_Click);
            // 
            // btnApplyChangesToDb
            // 
            this.btnApplyChangesToDb.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnApplyChangesToDb.Image = global::ExcelAddIn1.Properties.Resources.SaveSelectionToQuickTablesGallery;
            this.btnApplyChangesToDb.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnApplyChangesToDb.Name = "btnApplyChangesToDb";
            this.btnApplyChangesToDb.Size = new System.Drawing.Size(29, 24);
            this.btnApplyChangesToDb.Text = "Apply changes to Database";
            this.btnApplyChangesToDb.Click += new System.EventHandler(this.btnApplyChangesToDb_Click);
            // 
            // btnSaveChangesToFile
            // 
            this.btnSaveChangesToFile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnSaveChangesToFile.Image = global::ExcelAddIn1.Properties.Resources.SaveSelectionToTableOfContentsGallery;
            this.btnSaveChangesToFile.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSaveChangesToFile.Name = "btnSaveChangesToFile";
            this.btnSaveChangesToFile.Size = new System.Drawing.Size(29, 24);
            this.btnSaveChangesToFile.Text = "Save changes to File";
            this.btnSaveChangesToFile.Click += new System.EventHandler(this.btnSaveChangesToFile_Click);
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabDatabaseExplorer);
            this.tabMain.Controls.Add(this.tabPageSheetChanges);
            this.tabMain.Controls.Add(this.tabPageDataGeneration);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Location = new System.Drawing.Point(0, 27);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(322, 437);
            this.tabMain.TabIndex = 1;
            // 
            // tabDatabaseExplorer
            // 
            this.tabDatabaseExplorer.Controls.Add(this.tvTables);
            this.tabDatabaseExplorer.Location = new System.Drawing.Point(4, 25);
            this.tabDatabaseExplorer.Name = "tabDatabaseExplorer";
            this.tabDatabaseExplorer.Padding = new System.Windows.Forms.Padding(3);
            this.tabDatabaseExplorer.Size = new System.Drawing.Size(314, 408);
            this.tabDatabaseExplorer.TabIndex = 0;
            this.tabDatabaseExplorer.Text = "Database Explorer";
            this.tabDatabaseExplorer.UseVisualStyleBackColor = true;
            // 
            // tvTables
            // 
            this.tvTables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvTables.Location = new System.Drawing.Point(3, 3);
            this.tvTables.Name = "tvTables";
            this.tvTables.Size = new System.Drawing.Size(308, 402);
            this.tvTables.TabIndex = 2;
            this.tvTables.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvTables_NodeMouseDoubleClick);
            // 
            // tabPageSheetChanges
            // 
            this.tabPageSheetChanges.Controls.Add(this.lblRefresh);
            this.tabPageSheetChanges.Controls.Add(this.lvSheetChanges);
            this.tabPageSheetChanges.Location = new System.Drawing.Point(4, 25);
            this.tabPageSheetChanges.Name = "tabPageSheetChanges";
            this.tabPageSheetChanges.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSheetChanges.Size = new System.Drawing.Size(314, 408);
            this.tabPageSheetChanges.TabIndex = 1;
            this.tabPageSheetChanges.Text = "Sheet Changes";
            this.tabPageSheetChanges.UseVisualStyleBackColor = true;
            // 
            // lblRefresh
            // 
            this.lblRefresh.AutoSize = true;
            this.lblRefresh.Image = ((System.Drawing.Image)(resources.GetObject("lblRefresh.Image")));
            this.lblRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblRefresh.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lblRefresh.Location = new System.Drawing.Point(6, 3);
            this.lblRefresh.Name = "lblRefresh";
            this.lblRefresh.Size = new System.Drawing.Size(104, 17);
            this.lblRefresh.TabIndex = 2;
            this.lblRefresh.TabStop = true;
            this.lblRefresh.Text = "     Refresh List";
            this.lblRefresh.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblRefresh_LinkClicked);
            // 
            // lvSheetChanges
            // 
            this.lvSheetChanges.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvSheetChanges.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chPrimaryKey,
            this.chColName,
            this.chNewValue});
            this.lvSheetChanges.HideSelection = false;
            this.lvSheetChanges.Location = new System.Drawing.Point(0, 23);
            this.lvSheetChanges.Name = "lvSheetChanges";
            this.lvSheetChanges.Size = new System.Drawing.Size(311, 382);
            this.lvSheetChanges.TabIndex = 1;
            this.lvSheetChanges.UseCompatibleStateImageBehavior = false;
            this.lvSheetChanges.View = System.Windows.Forms.View.Details;
            this.lvSheetChanges.Visible = false;
            // 
            // chPrimaryKey
            // 
            this.chPrimaryKey.Text = "Primary Key";
            this.chPrimaryKey.Width = 80;
            // 
            // chColName
            // 
            this.chColName.Text = "Column Name";
            this.chColName.Width = 100;
            // 
            // chNewValue
            // 
            this.chNewValue.Text = "New Value";
            this.chNewValue.Width = 150;
            // 
            // tabPageDataGeneration
            // 
            this.tabPageDataGeneration.Controls.Add(this.btnInsertDataToSelection);
            this.tabPageDataGeneration.Controls.Add(this.cboColumnNames);
            this.tabPageDataGeneration.Controls.Add(this.label2);
            this.tabPageDataGeneration.Controls.Add(this.btnBrowseForDataFile);
            this.tabPageDataGeneration.Controls.Add(this.txtDataFile);
            this.tabPageDataGeneration.Controls.Add(this.label1);
            this.tabPageDataGeneration.Location = new System.Drawing.Point(4, 25);
            this.tabPageDataGeneration.Name = "tabPageDataGeneration";
            this.tabPageDataGeneration.Size = new System.Drawing.Size(314, 408);
            this.tabPageDataGeneration.TabIndex = 2;
            this.tabPageDataGeneration.Text = "Data Generation";
            this.tabPageDataGeneration.UseVisualStyleBackColor = true;
            // 
            // btnInsertDataToSelection
            // 
            this.btnInsertDataToSelection.Location = new System.Drawing.Point(82, 67);
            this.btnInsertDataToSelection.Name = "btnInsertDataToSelection";
            this.btnInsertDataToSelection.Size = new System.Drawing.Size(169, 23);
            this.btnInsertDataToSelection.TabIndex = 12;
            this.btnInsertDataToSelection.Text = "Insert random data in selection";
            this.btnInsertDataToSelection.UseVisualStyleBackColor = true;
            this.btnInsertDataToSelection.Click += new System.EventHandler(this.btnInsertDataToSelection_Click);
            // 
            // cboColumnNames
            // 
            this.cboColumnNames.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboColumnNames.FormattingEnabled = true;
            this.cboColumnNames.Location = new System.Drawing.Point(82, 39);
            this.cboColumnNames.Name = "cboColumnNames";
            this.cboColumnNames.Size = new System.Drawing.Size(185, 24);
            this.cboColumnNames.TabIndex = 11;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 17);
            this.label2.TabIndex = 10;
            this.label2.Text = "Data Column:";
            // 
            // btnBrowseForDataFile
            // 
            this.btnBrowseForDataFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseForDataFile.Location = new System.Drawing.Point(272, 11);
            this.btnBrowseForDataFile.Name = "btnBrowseForDataFile";
            this.btnBrowseForDataFile.Size = new System.Drawing.Size(24, 22);
            this.btnBrowseForDataFile.TabIndex = 9;
            this.btnBrowseForDataFile.Text = "...";
            this.btnBrowseForDataFile.UseVisualStyleBackColor = true;
            this.btnBrowseForDataFile.Click += new System.EventHandler(this.btnBrowseForDataFile_Click);
            // 
            // txtDataFile
            // 
            this.txtDataFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDataFile.Location = new System.Drawing.Point(82, 12);
            this.txtDataFile.Name = "txtDataFile";
            this.txtDataFile.Size = new System.Drawing.Size(185, 22);
            this.txtDataFile.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "Data file :";
            // 
            // diagOpenFile
            // 
            this.diagOpenFile.Filter = "CSV files|*.csv|All files|*.*";
            this.diagOpenFile.Title = "Select data file";
            // 
            // diagSaveFile
            // 
            this.diagSaveFile.DefaultExt = "sql";
            this.diagSaveFile.Filter = "SQL Files (*.sql)|*.sql|All files|*.*|Text files|*.txt";
            this.diagSaveFile.Title = "Save File As";
            // 
            // ExcelAddinTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabMain);
            this.Controls.Add(this.toolStrip1);
            this.Name = "ExcelAddinTaskPane";
            this.Size = new System.Drawing.Size(322, 464);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.tabMain.ResumeLayout(false);
            this.tabDatabaseExplorer.ResumeLayout(false);
            this.tabPageSheetChanges.ResumeLayout(false);
            this.tabPageSheetChanges.PerformLayout();
            this.tabPageDataGeneration.ResumeLayout(false);
            this.tabPageDataGeneration.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnConnectToDatabase;
        private System.Windows.Forms.ToolStripButton btnRefreshData;
        private System.Windows.Forms.ToolStripButton btnApplyChangesToDb;
        private System.Windows.Forms.ToolStripButton btnSaveChangesToFile;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabDatabaseExplorer;
        private System.Windows.Forms.TabPage tabPageSheetChanges;
        private System.Windows.Forms.LinkLabel lblRefresh;
        private System.Windows.Forms.ListView lvSheetChanges;
        private System.Windows.Forms.ColumnHeader chPrimaryKey;
        private System.Windows.Forms.ColumnHeader chColName;
        private System.Windows.Forms.ColumnHeader chNewValue;
        private System.Windows.Forms.TabPage tabPageDataGeneration;
        private System.Windows.Forms.Button btnInsertDataToSelection;
        private System.Windows.Forms.ComboBox cboColumnNames;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnBrowseForDataFile;
        private System.Windows.Forms.TextBox txtDataFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog diagOpenFile;
        private System.Windows.Forms.SaveFileDialog diagSaveFile;
        private System.Windows.Forms.TreeView tvTables;
    }
}
