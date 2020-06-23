using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
namespace SQLServerForExcel_Addin.Extensions
{
    public static class WorksheetExtensions
    {
        /// <summary>
        /// Checks whether the sheet has a primary key custom property and
        /// then return true or false indicating it is "connected" to a db table
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>bool</returns>
        public static bool ConnectedToDb(this Excel.Worksheet sheet)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;

            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    primaryKeyProperty = customProperties[i];
                    if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                }
                if (primaryKeyProperty != null)
                    return true;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return false;
        }

        public static string PrimaryKey(this Excel.Worksheet sheet)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;
            string keyName = null;

            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    primaryKeyProperty = customProperties[i];
                    if (primaryKeyProperty.Name == "PrimaryKey")
                    {
                        keyName = primaryKeyProperty.Value.ToString();
                    }
                    if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Marshal.ReleaseComObject(primaryKeyProperty);
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return keyName;
        }

        public static string ColumnName(this Excel.Worksheet sheet, int col)
        {
            string columnName = string.Empty;
            Excel.Range columnRange = null;

            try
            {
                string colLetter = ColumnIndexToColumnLetter(col);
                columnRange = sheet.Range[colLetter + "1:" + colLetter + "1"];
                if (columnRange != null)
                {
                    columnName = columnRange.Value.ToString();
                }
            }
            finally
            {
                if (columnRange != null) Marshal.ReleaseComObject(columnRange);
            }
            return columnName;
        }

        public static void AddChangedRow(this Excel.Worksheet sheet, int col, int row)
        {
            Excel.Range columnRange = null;
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range primaryKeyValueRange = null;
            Excel.Range rowValueRange = null;
            Excel.Range sheetCellRange = null;
            Excel.CustomProperty uncommittedChangesProperty = null;
            string primaryKey = string.Empty;
            string primaryKeyDataType = string.Empty;
            object primaryKeyValue = string.Empty;
            string columnName = string.Empty;
            object rowValue = string.Empty;
            string rowValueDataType = string.Empty;

            try
            {
                primaryKey = sheet.PrimaryKey();
                columnRange = sheet.Range["A1:CV1"];
                sheetCellRange = sheet.Cells;
                rowValueRange = sheetCellRange[row, col] as Excel.Range;
                primaryKeyColumnRange = columnRange.Find(primaryKey);

                if (primaryKeyColumnRange != null)
                {
                    primaryKeyValueRange = sheetCellRange[row, primaryKeyColumnRange.Column] as Excel.Range;
                    if (primaryKeyValueRange != null)
                    {
                        primaryKeyValue = primaryKeyValueRange.Value;
                        primaryKeyDataType = primaryKeyValue.GetType().ToString();
                    }
                }

                columnName = sheet.ColumnName(col);
                if (rowValueRange != null)
                {
                    rowValue = rowValueRange.Value;
                    rowValueDataType = rowValue.GetType().ToString();
                }

                string xmlString = "<row key=\"" + primaryKeyValue.ToString() + "\" ";
                xmlString += "keydatatype=\"" + primaryKeyDataType + "\" ";
                xmlString += "column=\"" + columnName + "\" ";
                xmlString += "columndatatype=\"" + rowValueDataType + "\">";
                xmlString += rowValue.ToString();
                xmlString += "</row>";
                xmlString = stripNonValidXMLCharacters(xmlString);

                uncommittedChangesProperty = sheet.GetProperty("UncommittedChanges");
                if (uncommittedChangesProperty == null)
                {
                    uncommittedChangesProperty = sheet.AddProperty("UncommittedChanges", xmlString);
                }
                else
                {
                    uncommittedChangesProperty.Value = uncommittedChangesProperty.Value + xmlString;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {

            }
        }

        public static void AddChangedRow(this Excel.Worksheet sheet, Excel.Range changedRange)
        {
            Excel.Range columnRange = null;
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range primaryKeyValueRange = null;
            Excel.Range rowValueRange = null;
            Excel.Range sheetCellsRange = null;
            Excel.Range rowsRange = null;
            Excel.Range colsRange = null;
            Excel.CustomProperty uncommittedChangesProperty = null;
            object rowValue = string.Empty;
            string rowValueDataType = string.Empty;
            string primaryKey = string.Empty;
            string primaryKeyDataType = string.Empty;
            object primaryKeyValue = string.Empty;            
            string columnName = string.Empty;
            string xmlString = string.Empty;

            try
            {
                primaryKey = sheet.PrimaryKey();
                columnRange = sheet.Range["A1:CV1"];
                sheetCellsRange = sheet.Cells;
                primaryKeyColumnRange = columnRange.Find(primaryKey, LookAt: Excel.XlLookAt.xlWhole);                
                rowsRange = changedRange.Rows;
                colsRange = rowsRange.Columns;                
                foreach (Excel.Range row in rowsRange)
                {
                    if (primaryKeyColumnRange != null)
                    {
                        int rowNum = row.Row;
                        int colNum = primaryKeyColumnRange.Column;
                        
                        primaryKeyValueRange = sheetCellsRange[rowNum, colNum] as Excel.Range;                        

                        if (primaryKeyValueRange != null)
                        {
                            try
                            {
                                primaryKeyValue = primaryKeyValueRange.Value;
                                primaryKeyDataType = primaryKeyValue.GetType().ToString();
                            }
                            catch (Exception)
                            {
                                primaryKeyValue = "";
                                primaryKeyDataType = "NULL";                                
                            }

                            foreach (Excel.Range col in colsRange)
                            {
                                colNum = col.Column;
                                columnName = sheet.ColumnName(colNum);
                                rowValueRange = sheetCellsRange[rowNum, col.Column] as Excel.Range;                                
                                if (rowValueRange != null)
                                {                                                                        
                                    try
                                    {
                                        rowValue = rowValueRange.Value;
                                        rowValueDataType = rowValue.GetType().ToString();
                                    }
                                    catch (System.NullReferenceException e)
                                    {
                                        rowValue = DBNull.Value;
                                        rowValueDataType = "null";                                        
                                    }
                                                                        
                                    xmlString += "<row key=\"" + primaryKeyValue.ToString() + "\" ";
                                    xmlString += "keydatatype=\"" + primaryKeyDataType + "\" ";
                                    xmlString += "column=\"" + columnName + "\" ";
                                    xmlString += "columndatatype=\"" + rowValueDataType + "\">";
                                    xmlString += rowValue.ToString();
                                    xmlString += "</row>";                                                                                                         
                                }
                            }
                        }
                        
                    }
                }

                uncommittedChangesProperty = sheet.GetProperty("UncommittedChanges");
                if (uncommittedChangesProperty == null)
                {
                    uncommittedChangesProperty = sheet.AddProperty("UncommittedChanges", xmlString);
                }
                else
                {
                    uncommittedChangesProperty.Value = uncommittedChangesProperty.Value + xmlString;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                //throw;
            }
            finally
            {
                if (uncommittedChangesProperty != null) Marshal.ReleaseComObject(uncommittedChangesProperty);
                if (colsRange != null) Marshal.ReleaseComObject(colsRange);
                if (rowsRange != null) Marshal.ReleaseComObject(rowsRange);
                if (sheetCellsRange != null) Marshal.ReleaseComObject(sheetCellsRange);
                if (rowValueRange != null) Marshal.ReleaseComObject(rowValueRange);
                if (primaryKeyValueRange != null) Marshal.ReleaseComObject(primaryKeyValueRange);
                if (primaryKeyColumnRange != null) Marshal.ReleaseComObject(primaryKeyColumnRange);
            }
        }

        public static Excel.CustomProperty AddProperty(this Excel.Worksheet sheet, string propertyName, object propertyValue)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty customProperty = null;

            try
            {
                customProperties = sheet.CustomProperties;
                customProperty = customProperties.Add(propertyName, propertyValue);
            }
            finally
            {
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return customProperty;
        }

        public static Excel.CustomProperty GetProperty(this Excel.Worksheet sheet, string propertyName)
        {
            Excel.CustomProperty customProperty = null;
            Excel.CustomProperties customProperties = null;
            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    customProperty = customProperties[i];                    
                    if (customProperty != null && customProperty.Name.ToLower() == propertyName.ToLower())
                    {
                        return customProperty;
                    }
                    else
                    {
                        customProperty = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (customProperties != null) Marshal.ReleaseComObject(customProperties);
            }
            return customProperty;
        }

        public static string ChangesToSql(this Excel.Worksheet sheet, string tableName, string primaryKeyName)
        {
            Excel.CustomProperty customProperty = null;
            string xml = string.Empty;
            string sql = string.Empty;

            try
            {
                customProperty = sheet.GetProperty("uncommittedchanges");
                if (customProperty != null)
                {
                    xml = ToSafeXml("<uncommittedchanges>" + customProperty.Value.ToString() + "</uncommittedchanges>");
                    XDocument doc = XDocument.Parse(xml);                    
                    foreach (var dm in doc.Descendants("row"))
                    {

                        string key = dm.Attribute("key").Value;
                        string keyDataType = dm.Attribute("keydatatype").Value;
                        string column = dm.Attribute("column").Value;
                        string columnDataType = dm.Attribute("columndatatype").Value;
                        string value = dm.Value;

                        if (keyDataType != "NULL")
                        {
                            sql += "UPDATE " + tableName + " SET " + column + " = ";

                            if (columnDataType.ToLower().Contains("date") || columnDataType.ToLower().Contains("string") || columnDataType.ToLower().Contains("boolean"))
                            {
                                sql += "'" + value + "'";
                            }
                            else if (columnDataType.ToLower().Contains("null"))
                            {
                                sql += "null";
                            }
                            else
                            {
                                sql += value;
                            }

                            sql += " WHERE " + primaryKeyName + " = ";

                            if (keyDataType.ToLower().Contains("date") || keyDataType.ToLower().Contains("string"))
                            {
                                sql += "'" + key + "'";
                            }
                            else
                            {
                                sql += key;
                            }

                            sql += Environment.NewLine;
                        }                                              
                    }
                }
            }
            finally
            {
                if (customProperty != null) Marshal.ReleaseComObject(customProperty);
            }
            return sql;
        }

        private static string ToSafeXml(string xmlString)
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

        private static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        private static String stripNonValidXMLCharacters(string textIn)
        {
            StringBuilder textOut = new StringBuilder(); // Used to hold the output.
            char current; // Used to reference the current character.


            if (textIn == null || textIn == string.Empty) return string.Empty; // vacancy test.
            for (int i = 0; i < textIn.Length; i++)
            {
                current = textIn[i];


                if ((current == 0x9 || current == 0xA || current == 0xD) ||
                    ((current >= 0x20) && (current <= 0xD7FF)) ||
                    ((current >= 0xE000) && (current <= 0xFFFD)) ||
                    ((current >= 0x10000) && (current <= 0x10FFFF)))
                {
                    textOut.Append(current);
                }
            }
            return textOut.ToString();
        }

        public static string DeleteRowsFromTable(this Excel.Worksheet sheet, string tableName, bool refresh)
        {
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range columnRange = null;
            Excel.Range sheetCellsRange = null;
            string primaryKey = string.Empty;
            string sql = string.Empty;
            Excel.Range deletedRows = null;
            double deletedRowNum;
            bool rowsDeleted = false;
           
            try
            {                
                deletedRows = sheet.Range["A:A"].SpecialCells(Excel.XlCellType.xlCellTypeBlanks).EntireRow;
                //got at least 1 empty row
                foreach (Excel.Range row in deletedRows)
                {
                    deletedRowNum = sheet.Application.WorksheetFunction.CountA(row);
                    if (deletedRowNum == 0)
                    {
                        rowsDeleted = true;
                    }
                }
                
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                rowsDeleted = false;
            }

            if (rowsDeleted == true)
            {

                primaryKey = sheet.PrimaryKey();
                columnRange = sheet.Range["A1:CV1"];
                sheetCellsRange = sheet.Cells;
                primaryKeyColumnRange = columnRange.Find(primaryKey, LookAt: Excel.XlLookAt.xlWhole);

                object[,] pkValues = (object[,])sheet.Columns[primaryKeyColumnRange.Column].Cells.Value;

                List<string> primaryKeyValues = pkValues.Cast<object>().ToList().ConvertAll(x => Convert.ToString(x));
                primaryKeyValues.RemoveAt(0);
                primaryKeyValues.RemoveAll(str => String.IsNullOrEmpty(str));

                string primaryKeyValuesJoined = string.Join(",", primaryKeyValues);
                primaryKeyValuesJoined = "'" + primaryKeyValuesJoined.Replace(",", "','") + "'";

                if (refresh == false)
                {
                    sql = "Delete from " + tableName + " Where " + primaryKey + " NOT IN( " + primaryKeyValuesJoined + ")";
                }
                else if (refresh == true)
                {
                    sql = "Select "+ primaryKey + " from " + tableName + " Where " + primaryKey + " NOT IN( " + primaryKeyValuesJoined + ")";
                }

            }
            //Debug.WriteLine(sql);

            return sql;

        }

        public static string InsertRowsIntoTable(this Excel.Worksheet sheet, string tableName)
        {
            string sql = string.Empty;
            List<string> rowValues = new List<string>();
            string rowValuesJoined = string.Empty;
            string primaryKey = string.Empty;
            Excel.CustomProperty tableColumnsProperty = null;
            Dictionary<string,string> tableColumnTypes = new Dictionary<string, string>();
            string xml = string.Empty;

            tableColumnsProperty = sheet.GetProperty("TableColumns");

            try
            {
                if (tableColumnsProperty != null)
                {
                    xml = ToSafeXml("<tablecolumns>" + tableColumnsProperty.Value.ToString() + "</tablecolumns>");
                    XDocument doc = XDocument.Parse(xml);                    
                    foreach (var dm in doc.Descendants("row"))
                    {                        
                        string colValue = dm.Attribute("column").Value;
                        string colDataTypeValue = dm.Attribute("columndatatype").Value;

                        tableColumnTypes.Add(colValue, colDataTypeValue);
                    }

                }
            }
            catch (System.NullReferenceException e)
            {
                Console.Write(e.Message);             
            }                                 

            primaryKey = sheet.PrimaryKey();
            //int lastTableRow;
            int lastTableColumn;

            //https://stackoverflow.com/questions/7674573/programmatically-getting-the-last-filled-excel-row-using-c-sharp
            // Find the last real row
            //lastTableRow = sheet.Cells.Find("*", System.Reflection.Missing.Value,System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            lastTableColumn = sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            
            try
            {
                Excel.Range insertRows = sheet.Columns["A:A"].SpecialCells(Excel.XlCellType.xlCellTypeBlanks).EntireRow as Excel.Range;

                Excel.Range headerRow = sheet.Rows["1:1"].EntireRow as Excel.Range;

                object[,] hdrValues = (object[,])headerRow.Cells.Value;

                List<string> headerRowValues = hdrValues.Cast<object>().ToList().ConvertAll(x => Convert.ToString(x));
                headerRowValues.RemoveAll(header => header == string.Empty);
                headerRowValues.Remove(primaryKey);
                string headerRowValuesJoined = string.Join(",", headerRowValues);

                foreach (Excel.Range row in insertRows)
                {
                    object[,] rwValues = (object[,])row.Cells.Value2;                   
                    rowValues = rwValues.Cast<object>().ToList().ConvertAll(x => Convert.ToString(x));                    
                    rowValues.RemoveRange(lastTableColumn, rowValues.Count - lastTableColumn);
                    rowValues.RemoveAt(0);                                       

                    for (int i = 0; i < rowValues.Count; i++)
                    {
                        try
                        {
                            if (tableColumnTypes[headerRowValues[i]].ToString().Contains("date") & !string.IsNullOrEmpty(rowValues[i]))
                            {
                                rowValues[i] = DateTime.FromOADate(Convert.ToDouble(rowValues[i])).ToString(CultureInfo.InvariantCulture);
                            }
                        }
                        catch (System.Collections.Generic.KeyNotFoundException e)
                        {
                            Console.WriteLine(e.Message);
                        }                                               

                        if (string.IsNullOrEmpty(rowValues[i]))
                        {
                            rowValues[i] = "NULL";
                        }
                    }

                    rowValuesJoined = string.Join(",", rowValues);
                    rowValuesJoined = "'" + rowValuesJoined.Replace(",", "','") + "'";
                    rowValuesJoined = rowValuesJoined.Replace("'NULL'", "NULL");

                    HashSet<string> unique_items = new HashSet<string>(rowValuesJoined.Split(','));
                    
                    if (unique_items.Count != 1)
                    {
                        sql += "Insert into " + tableName + "(" + headerRowValuesJoined + ") VALUES( " + rowValuesJoined + ")";
                        sql += Environment.NewLine;
                    }
                    
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.Write(e.Message);
            }
            

            //Debug.WriteLine(sql);

            return sql;
        }



    }
}
