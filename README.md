# EditSQLTableInExcel

## Import, edit, insert and delete data for SQL Server in Excel.

EditSQLTableInExcel is an Excel Add-in that provides users a way to import and edit table data in Excel.
It allows users to insert, update and delete records in any SQL Server table that has a Primary Key defined.
It has been designed to be a simple tool for data analysts to manipulate SQL data without having to write scripts.

This Add-in is heavily based on Pieter van der Westhuizen's (Pietervdw) SQL Server for Excel Add-in which is available at [https://github.com/Pietervdw/SQLForExcel](https://github.com/Pietervdw/SQLForExcel).
This version fixes multiple bugs found in the original Add-in, and no longer needs [Add-in Express for Office and .net](https://www.add-in-express.com/add-in-net/index.php) in order to build and run the project.


## Setup

To install the software, download the .msi file from the latest release from the releases tab and follow the installation instructions.
After the installation has finished the Add-in should appear in Excel in the Data tab.
Make sure to reopen any Excel files that were open during the install process to see the Add-in.
If by any chance it doesn't appear check the COM Add-ins menu in the Developer tab in Excel and make sure it is checked.
