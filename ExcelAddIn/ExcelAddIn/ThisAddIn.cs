﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using SQLServerForExcel_Addin;
using System.Diagnostics;
using SQLServerForExcel_Addin.Extensions;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private ExcelAddinTaskPane taskPaneControl1;
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;       
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {            
            taskPaneControl1 = new ExcelAddinTaskPane();
            taskPaneValue = this.CustomTaskPanes.Add(
                taskPaneControl1, "SQL Server For Excel");
            taskPaneValue.VisibleChanged +=
                new EventHandler(taskPaneValue_VisibleChanged);
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
            Excel.Worksheet activeSheet = ((Excel.Worksheet)Application.ActiveSheet);
            activeSheet.AddChangedRow(Target);           
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
