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
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabPageInstructions = new System.Windows.Forms.TabPage();
            this.textBoxInstructions = new System.Windows.Forms.TextBox();
            this.tabDatabaseExplorer = new System.Windows.Forms.TabPage();
            this.tvTables = new System.Windows.Forms.TreeView();
            this.tabPageSheetChanges = new System.Windows.Forms.TabPage();
            this.lblRefresh = new System.Windows.Forms.LinkLabel();
            this.lvSheetChanges = new System.Windows.Forms.ListView();
            this.chPrimaryKey = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chColName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chNewValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.diagOpenFile = new System.Windows.Forms.OpenFileDialog();
            this.diagSaveFile = new System.Windows.Forms.SaveFileDialog();
            this.toolStrip1.SuspendLayout();
            this.tabMain.SuspendLayout();
            this.tabPageInstructions.SuspendLayout();
            this.tabDatabaseExplorer.SuspendLayout();
            this.tabPageSheetChanges.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnConnectToDatabase,
            this.btnRefreshData,
            this.btnApplyChangesToDb});
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
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabPageInstructions);
            this.tabMain.Controls.Add(this.tabDatabaseExplorer);
            this.tabMain.Controls.Add(this.tabPageSheetChanges);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Location = new System.Drawing.Point(0, 27);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(322, 437);
            this.tabMain.TabIndex = 1;
            // 
            // tabPageInstructions
            // 
            this.tabPageInstructions.Controls.Add(this.textBoxInstructions);
            this.tabPageInstructions.Location = new System.Drawing.Point(4, 25);
            this.tabPageInstructions.Name = "tabPageInstructions";
            this.tabPageInstructions.Size = new System.Drawing.Size(314, 408);
            this.tabPageInstructions.TabIndex = 2;
            this.tabPageInstructions.Text = "Instructions";
            this.tabPageInstructions.UseVisualStyleBackColor = true;
            // 
            // textBoxInstructions
            // 
            this.textBoxInstructions.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxInstructions.Location = new System.Drawing.Point(3, 3);
            this.textBoxInstructions.Multiline = true;
            this.textBoxInstructions.Name = "textBoxInstructions";
            this.textBoxInstructions.Size = new System.Drawing.Size(308, 402);
            this.textBoxInstructions.TabIndex = 0;
            this.textBoxInstructions.Text = resources.GetString("textBoxInstructions.Text");
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
            this.tabPageInstructions.ResumeLayout(false);
            this.tabPageInstructions.PerformLayout();
            this.tabDatabaseExplorer.ResumeLayout(false);
            this.tabPageSheetChanges.ResumeLayout(false);
            this.tabPageSheetChanges.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnConnectToDatabase;
        private System.Windows.Forms.ToolStripButton btnRefreshData;
        private System.Windows.Forms.ToolStripButton btnApplyChangesToDb;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabDatabaseExplorer;
        private System.Windows.Forms.TabPage tabPageSheetChanges;
        private System.Windows.Forms.LinkLabel lblRefresh;
        private System.Windows.Forms.ListView lvSheetChanges;
        private System.Windows.Forms.ColumnHeader chPrimaryKey;
        private System.Windows.Forms.ColumnHeader chColName;
        private System.Windows.Forms.ColumnHeader chNewValue;
        private System.Windows.Forms.OpenFileDialog diagOpenFile;
        private System.Windows.Forms.SaveFileDialog diagSaveFile;
        private System.Windows.Forms.TreeView tvTables;
        private System.Windows.Forms.TabPage tabPageInstructions;
        private System.Windows.Forms.TextBox textBoxInstructions;
    }
}
