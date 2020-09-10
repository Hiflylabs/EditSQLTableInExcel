using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using SQLServerForExcel_Addin;
using SQLServerForExcel_Addin.Extensions;
using Microsoft.Office.Tools;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {        
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;
        public Dictionary<string, CustomTaskPane> ExcelCustomTaskPanes = new Dictionary<string, CustomTaskPane>();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {            
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            this.Application.SheetChange += Application_SheetChange;            

        }

        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            Globals.Ribbons.Ribbon1.toggleButton1.Checked =
                taskPaneValue.Visible;
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return taskPaneValue;
            }
        }

    private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            Excel.CustomProperty tableLoadedProperty = null;
            Excel.Worksheet activeSheet = ((Excel.Worksheet)Application.ActiveSheet);
            tableLoadedProperty = activeSheet.GetProperty("TableLoaded");
            if (tableLoadedProperty != null)
            {
                activeSheet.AddChangedRow(Target);
            }                
        }

    private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            var wbCtp = ExcelCustomTaskPanes.Where(wb => wb.Key == Wb.FullName).FirstOrDefault().Value;
            if (wbCtp == null)
            {
                Globals.Ribbons.Ribbon1.toggleButton1.Checked = false;
                taskPaneValue = this.CustomTaskPanes.Add(new ExcelAddinTaskPane(),
                "SQL Server For Excel", Wb.Windows[1]);
                taskPaneValue.VisibleChanged +=
                    new EventHandler(taskPaneValue_VisibleChanged);
                ExcelCustomTaskPanes.Add(Wb.FullName, taskPaneValue);
            }
            else
            {
                taskPaneValue = wbCtp;
                if (taskPaneValue.Visible)
                {
                    Globals.Ribbons.Ribbon1.toggleButton1.Checked = true;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.toggleButton1.Checked = false;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
