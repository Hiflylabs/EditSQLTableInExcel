using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using ExcelAddIn1;
using GenericParsing;
using Microsoft.Data.ConnectionUI;
using Microsoft.Office.Core;
using SQLServerForExcel_Addin.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace SQLServerForExcel_Addin
{
    public partial class ExcelAddinTaskPane : UserControl
    {
        DataConnectionDialog dcd;
        string dbName = string.Empty;
        string tableName = string.Empty;
        string serverName = string.Empty;
        string connectionString = string.Empty;
        private System.Data.DataTable sourceData = null;       

        public Excel.Application ExcelApp = Globals.ThisAddIn.Application as Excel.Application;

        public ExcelAddinTaskPane()
        {
            InitializeComponent();
            dcd = new DataConnectionDialog();
        }

        private void btnConnectToDatabase_Click(object sender, EventArgs e)
        {
            
            DataConnectionConfiguration dcs = new DataConnectionConfiguration(null);
            dcs.LoadConfiguration(dcd);

            if (DataConnectionDialog.Show(dcd) == DialogResult.OK)
            {
                var tables = SqlUtils.GetAllTables(dcd.ConnectionString);

                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.ConnectionString = dcd.ConnectionString;
                dbName = builder.InitialCatalog;
                serverName = builder.DataSource;
                connectionString = dcd.ConnectionString;

                TreeNode rootNode = new TreeNode(builder.InitialCatalog, 1, 1);
                TreeNode tablesNode = rootNode.Nodes.Add("Tables", "Tables", 2, 2);
                tablesNode.Tag = dcd.ConnectionString;

                foreach (string table in tables)
                {
                    TreeNode tableNode = tablesNode.Nodes.Add(table, table, 3, 3);
                    tableNode.Tag = "table";
                }

                tvTables.Nodes.Clear();
                tvTables.Nodes.Add(rootNode);
                tvTables.ExpandAll();
            }
        }

        private void tvTables_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            Excel.Worksheet sheet = null;
            Excel.Range insertionRange = null;
            Excel.QueryTable queryTable = null;
            Excel.QueryTables queryTables = null;
            Excel.Range cellRange = null;
            Excel.CustomProperties sheetProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;
            Excel.CustomProperty tableColumnsProperty = null;

            SqlConnectionStringBuilder builder = null;
            string connString = "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=@servername;Initial Catalog=@databasename";
            string connStringSQL = "OLEDB;Provider=SQLOLEDB.1;Persist Security Info=True;User ID=@username;Password=@password;Data Source=@servername;Initial Catalog=@databasename";
            string databaseName = string.Empty;
            string tableName = string.Empty;
            string xmlString = string.Empty;

            try
            {
                ExcelApp.EnableEvents = false;
                tableName = e.Node.Text;                
                sheet = ExcelApp.ActiveSheet as Excel.Worksheet;                
                cellRange = sheet.Cells;
                insertionRange = cellRange[1, 1] as Excel.Range;                
                builder = new SqlConnectionStringBuilder(dcd.ConnectionString);
                databaseName = builder.InitialCatalog;                
                if (!builder.IntegratedSecurity)
                    connString = connStringSQL;                
                connString =
                    connString.Replace("@servername", builder.DataSource)
                        .Replace("@databasename", databaseName)
                        .Replace("@username", builder.UserID)
                        .Replace("@password", builder.Password);
                queryTables = sheet.QueryTables;                
                if (queryTables.Count > 0)
                {
                    queryTable = queryTables.Item(1);
                    queryTable.CommandText = String.Format("SELECT * FROM [{0}].{1}", databaseName, tableName);
                }
                else
                {
                    queryTable = queryTables.Add(connString, insertionRange,
                        String.Format("SELECT * FROM [{0}].{1}", databaseName, tableName));
                }
                queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells;
                queryTable.PreserveColumnInfo = true;
                queryTable.PreserveFormatting = true;
                queryTable.Refresh(false);                
                var primaryKey = SqlUtils.GetPrimaryKey(dcd.ConnectionString, tableName);
                var tableColumns = SqlUtils.GetAllColumns(dcd.ConnectionString, tableName);

                // save original table 
                this.tableName = tableName;
                // to sheet name must be less then 31 characters long 
                sheet.Name = tableName.Substring(0, Math.Min(tableName.Length, 30));

                chPrimaryKey.Text = primaryKey;

                sheetProperties = sheet.CustomProperties;
                primaryKeyProperty = sheetProperties.Add("PrimaryKey", primaryKey);

                foreach (var cols in tableColumns)
                {
                    xmlString += "<row column=\"" + cols.Key + "\" ";
                    xmlString += "columndatatype=\"" + cols.Value + "\">";                    
                    xmlString += cols.Key;
                    xmlString += "</row>";

                }
                                
                tableColumnsProperty = sheetProperties.Add("TableColumns", xmlString);
                ExcelApp.EnableEvents = true;
        }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                throw;
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (tableColumnsProperty != null) Marshal.ReleaseComObject(tableColumnsProperty);
                if (sheetProperties != null) Marshal.ReleaseComObject(sheetProperties);
                if (cellRange != null) Marshal.ReleaseComObject(cellRange);
                if (queryTables != null) Marshal.ReleaseComObject(queryTables);
                if (queryTable != null) Marshal.ReleaseComObject(queryTable);
                if (insertionRange != null) Marshal.ReleaseComObject(insertionRange);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        public void RefreshChanges()
        {
            Excel.Worksheet activeSheet = null;
            Excel.CustomProperty changesProperty = null;
            string xml = string.Empty;
            string sql = string.Empty;

            try
            {
                activeSheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                changesProperty = activeSheet.GetProperty("UncommittedChanges");
                lvSheetChanges.Items.Clear();                
                if (changesProperty != null)
                {
                    lvSheetChanges.Visible = true;
                    xml = ToSafeXml("<uncommittedchanges>" + changesProperty.Value.ToString() + "</uncommittedchanges>");
                    XDocument doc = XDocument.Parse(xml);
                    foreach (var dm in doc.Descendants("row"))
                    {
                        ListViewItem item = new ListViewItem(new string[]
                        {
                            dm.Attribute("key").Value,
                            dm.Attribute("column").Value,
                            dm.Value
                        });
                        lvSheetChanges.Items.Add(item);
                    }
                }

                sql = activeSheet.DeleteRowsFromTable(this.tableName,true);
                var primaryKey = activeSheet.PrimaryKey();
                if (!string.IsNullOrEmpty(sql))
                {
                    using (SqlConnection conn = new SqlConnection(dcd.ConnectionString))
                    {
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader dbReader;
                            conn.Open();
                            dbReader = cmd.ExecuteReader();
                            while (dbReader.Read())
                            {
                                ListViewItem item = new ListViewItem(new string[]
                                {
                                    dbReader[primaryKey].ToString(),
                                    primaryKey,
                                    "delete"
                                });
                                lvSheetChanges.Items.Add(item);                                
                            }
                        }
                        conn.Close();
                    }                    
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);                
            }
            finally
            {
                //if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
                //if (changesProperty != null) Marshal.ReleaseComObject(changesProperty);
            }
        }            

        private void lblRefresh_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RefreshChanges();
        }

        public static string ToSafeXml(string xmlString)
        {
            try
            {
                if ((xmlString != null))
                {
                    xmlString = xmlString.Replace("&", "&amp;");
                    xmlString = xmlString.Replace("'", "''");
                    //xmlString = xmlString.Replace(">", "&gt;");
                    //xmlString = xmlString.Replace("<", "&lt;");
                    //xmlString = xmlString.Replace("\"", "&quot;");
                    xmlString = xmlString.Replace("–", "-");
                    return xmlString;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception Er)
            {
                return "";
            }
        }              

        private void RefreshSheetData()
        {
            Excel.Worksheet sheet = null;
            Excel.QueryTables queryTables = null;
            Excel.QueryTable queryTable = null;
            Excel.CustomProperty changesProperty = null;

            try
            {
                ExcelApp.EnableEvents = false;
                sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                sheet.Cells.ClearContents();
                if (sheet != null)
                {
                    queryTables = sheet.QueryTables;

                    if (queryTables.Count > 0)
                    {
                        queryTable = queryTables.Item(1);
                        queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells;
                        queryTable.PreserveColumnInfo = true;
                        queryTable.PreserveFormatting = true;
                        queryTable.Refresh(false);
                    }
                    changesProperty = sheet.GetProperty("uncommittedchanges");
                    if (changesProperty != null)
                        changesProperty.Delete();
                }
                ExcelApp.EnableEvents = true;
            }
            finally
            {
                if (changesProperty != null) Marshal.ReleaseComObject(changesProperty);
                if (queryTable != null) Marshal.ReleaseComObject(queryTable);
                if (queryTables != null) Marshal.ReleaseComObject(queryTables);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }

        private void btnRefreshData_Click(object sender, EventArgs e)
        {
            RefreshSheetData();
        }

        private void btnApplyChangesToDb_Click(object sender, EventArgs e)
        {

            Excel.Worksheet sheet = null;
            Excel.CustomProperty primaryKeyProperty = null;
            string primaryKey = string.Empty;

            try
            {
                if (MessageBox.Show("This will commit the changes to the database. This action cannot be reversed. Are you sure?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                    if (sheet != null)
                    {
                        primaryKeyProperty = sheet.GetProperty("PrimaryKey");
                        if (primaryKeyProperty != null)
                        {
                            primaryKey = primaryKeyProperty.Value.ToString();
                            string sql = sheet.ChangesToSql(this.tableName, primaryKey);
                            sql += Environment.NewLine;
                            sql += sheet.DeleteRowsFromTable(this.tableName,false);
                            sql += Environment.NewLine;
                            sql += sheet.InsertRowsIntoTable(this.tableName);

                            if (!string.IsNullOrEmpty(sql))
                            {
                                using (SqlConnection conn = new SqlConnection(dcd.ConnectionString))
                                {
                                    SqlCommand cmd = new SqlCommand(sql, conn);
                                    if (conn.State == ConnectionState.Closed)
                                    {
                                        conn.Open();
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                RefreshSheetData();
                            }
                        }
                    }
                }
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }

        }

        //http://www.codeproject.com/Articles/11698/A-Portable-and-Efficient-Generic-Parser-for-Flat-F
    }
}
