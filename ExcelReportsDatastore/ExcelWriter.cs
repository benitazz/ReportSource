#region

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.IO;
using System.Linq;

using ExcelReportsDatastore.Helpers;
using ExcelReportsDatastore.Reports.Data;

using ExcelReportsUtils;
using ExcelReportsUtils.Extensions;

using Microsoft.Office.Interop.Excel;

using DataTable = System.Data.DataTable;

#endregion

namespace ExcelReportsDatastore
{
    /// <summary>
    /// Database to excel writer.
    /// </summary>
    public class ExcelWriter
    {
        #region Member Variable

        /// <summary>
        /// The background worker.
        /// </summary>
        private static BackgroundWorker backgroundWorker;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The dissect records.
        /// </summary>
        /// <param name="columnName">
        /// The column name.
        /// </param>
        /// <param name="tableName">
        /// The table name.
        /// </param>
        /// <param name="fileName">
        /// The actual file name.
        /// </param>
        /// <param name="directory">
        /// the directory.
        /// </param>
        /// <param name="background">
        /// The background.
        /// </param>
        public static void DissectRecords(IReportData dissectHelper, BackgroundWorker background)
        {
            if (dissectHelper == null)
            {
                return;
            }

            if (string.IsNullOrEmpty(dissectHelper.ColumnNameFilter))
            {
                return;
            }

            backgroundWorker = background;

           backgroundWorker.RunWorkerAsync(dissectHelper);       
        }

        /// <summary>
        /// Dissects the specified dissect helper.
        /// </summary>
        /// <param name="reportData">The report data.</param>
        public static void Dissect(IReportData reportData)
        {
            var connection = ExcelOleDbReader.GetLocalConnection();
            try
            {
               if (reportData == null)
                {
                    return;
                }

                connection.Open();

                var selectDisctinctDissectQuery = string.Format(
                    "select DISTINCT {0} from {1}",
                    reportData.ColumnNameFilter,
                    "ExcelDataTable");

                using (var cmd = new SqlCeCommand(selectDisctinctDissectQuery, connection))
                {
                    using (var dissectColumnReader = cmd.ExecuteReader())
                    {
                        while (dissectColumnReader.Read())
                        {
                            string dissectValue = dissectColumnReader.GetString(0);

                            if (string.IsNullOrEmpty(dissectValue))
                            {
                                continue;
                            }

                            var newValue = dissectValue.Replace("'", "''");

                            var dissectQuery = string.Format(
                                "select * from {0} Where [{1}]='{2}'",
                                "ExcelDataTable",
                                reportData.ColumnNameFilter,
                                newValue);

                            var dataTable = DatabaseReader.GetDatabaseDataTable(connection, dissectQuery);

                            dissectValue = dissectValue.RemoveInvalidCharacterForFilename();

                            var filename = string.Format(
                                @"{0}\{1}_{2}",
                                reportData.Directory,
                                dissectValue,
                                Path.GetFileName(reportData.FileName));

                            var sheetName = dissectValue;

                            // var filename = string.Format(@"{0}\{1}.{2}", directory, dissectValue, Path.GetExtension(sheetName));
                            if (sheetName.Length > 30)
                            {
                                sheetName = sheetName.Substring(0, 30);
                            }

                            if (dataTable.Rows.Count > 10000)
                            {
                                ExportBigDataToXlsx(dataTable, filename, sheetName);
                            }
                            else
                            {
                                ExportToXlsx(dataTable, filename, sheetName);
                            }
                        }
                    }
                }

                // run the rest of the program
            }
            catch (Exception ex)
            {
               ExcelReportsUtils.Dialogs.ShowError(ex);
            }
            finally
            {
                connection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Tests the export1.
        /// </summary>
        /// <param name="table">
        /// The table.
        /// </param>
        /// <param name="fileName">
        /// Name of the file.
        /// </param>
        /// <param name="sheetName">
        /// Name of the sheetName.
        /// </param>
        public static void ExportBigDataToXlsx(DataTable table, string fileName, string sheetName)
        {
            // Check if there are rows to process
            if (table != null && table.Rows.Count > 0)
            {
                // Determine the number of chunks
                int chunkSize = 100000;
                double chunkCountD = table.Rows.Count / (double)chunkSize;
                int chunkCount = table.Rows.Count / chunkSize;
                chunkCount = chunkCountD > chunkCount ? chunkCount + 1 : chunkCount;

                // Instantiate excel
                var excel = new Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };

                // Get a workbook
                Workbook wb = excel.Workbooks.Add();

                // Get a worksheet
                Worksheet ws = wb.Worksheets.Add();
                ws.Name = sheetName;

                // Add column names to excel
                int col = 1;
                foreach (DataColumn c in table.Columns)
                {
                    ws.Cells[1, col] = c.ColumnName;
                    col++;
                }

                // Build 2D array
                int i = 0;
                var data = new string[table.Rows.Count, table.Columns.Count];

                foreach (DataRow row in table.Rows)
                {
                    int j = 0;

                    foreach (DataColumn c in table.Columns)
                    {
                        data[i, j] = row[c].ToString();
                        j++;
                    }

                    i++;
                }

                int processed = 0;
                int data2DLength = data.GetLength(1);

                for (int chunk = 1; chunk <= chunkCount; chunk++)
                {
                    if (table.Rows.Count - processed < chunkSize)
                    {
                        chunkSize = table.Rows.Count - processed;
                    }

                    var chunkData = new string[chunkSize, data2DLength];
                    int l = 0;

                    for (int k = processed; k < chunkSize + processed; k++)
                    {
                        for (int m = 0; m < data2DLength; m++)
                        {
                            chunkData[l, m] = table.Rows[k][m].ToString();
                        }

                        l++;
                    }

                    // Set the range value to the chunk 2d array
                    ws.Range[ws.Cells[2 + processed, 1], ws.Cells[processed + chunkSize + 1, data2DLength]].Value2 =
                        chunkData;
                    processed += chunkSize;
                }

                wb.SaveAs(
                    fileName, 
                    XlFileFormat.xlWorkbookDefault, 
                    Type.Missing, 
                    Type.Missing, 
                    false, 
                    false, 
                    XlSaveAsAccessMode.xlNoChange, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing);

                wb.Close();
                excel.Quit();
            }
        }

        /// <summary>
        /// Exports to XLSX.
        /// </summary>
        /// <param name="dataTableToExport">
        /// The data table to export.
        /// </param>
        /// <param name="fileName">
        /// Name of the table.
        /// </param>
        /// <param name="sheetToCreate">
        /// The sheetName to create.
        /// </param>
        public static void ExportToXlsx(DataTable dataTableToExport, string fileName, string sheetToCreate)
        {
            var rows = dataTableToExport.Rows.Cast<DataRow>().ToList();

            ExportToXlsx(rows, dataTableToExport, fileName, sheetToCreate);
        }

        /// <summary>
        /// Exports to XLSX.
        /// </summary>
        /// <param name="selectedRows">The selected rows.</param>
        /// <param name="dataTable">The data table.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        public static void ExportToXlsx(
            List<DataRow> selectedRows, 
            DataTable dataTable,
             string fileName,
            string sheetName)
        {
            const char Space = ' ';
            string dest = fileName;

            if (File.Exists(dest))
            {
                File.Delete(dest);
            }

            fileName = dest;

            // fileName = fileName.ReplaceSpecialCharaterWithUnderScore();

            if (sheetName == null)
            {
                sheetName = string.Empty;
            }

            sheetName = sheetName.Trim().Replace(Space, '_');

            if (sheetName.Length == 0)
            {
                sheetName = dataTable.TableName.Replace(Space, '_');
            }

            if (sheetName.Length == 0)
            {
                sheetName = "NoTableName";
            }

            sheetName = sheetName.ReplaceSpecialCharaterWithUnderScore();

            if (sheetName.Length > 30)
            {
                sheetName = sheetName.Substring(0, 30);
            }

            // Excel names are less than 31 chars
            string queryCreateExcelTable = "CREATE TABLE [" + sheetName + "] (";
            var colNames = new Dictionary<string, string>();

            foreach (DataColumn dc in dataTable.Columns)
            {
                // Cause the query to name each of the columns to be created.
                string modifiedcolName = dc.ColumnName; // .Replace(Space, '_').Replace('.', '#');
                string origColName = dc.ColumnName;
                colNames.Add(modifiedcolName, origColName);

                switch (dc.DataType.ToString())
                {
                    case "System.String":
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " text,";
                        break;
                    case "System.DateTime":
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " datetime,";
                        break;
                    case "System.Boolean":
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " LOGICAL,";
                        break;
                    case "System.Byte":
                    case "System.Int16":
                    case "System.Int32":
                    case "System.Int64":
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " int,";
                        break;
                    case "System.Decimal":
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " decimal,";
                        break;
                    case "System.Double":
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " double,";
                        break;
                    default:
                        queryCreateExcelTable += "[" + modifiedcolName + "]" + " text,";
                        break;
                }
            }

            queryCreateExcelTable = queryCreateExcelTable.TrimEnd(new[] { Convert.ToChar(",") }) + ")";

            // adds the closing parentheses to the query string
            if (selectedRows.Count > 65000 && fileName.ToLower().EndsWith(".xls"))
            {
                // use Excel 2007 for large sheets.
                fileName = fileName.ToLower().Replace(".xls", string.Empty) + ".xlsx";
            }

            string strCn = string.Empty;
            var extension = Path.GetExtension(fileName);

            if (extension != null)
            {
                string ext = extension.ToLower();

                if (ext == ".xls")
                {
                    strCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName
                            + "; Extended Properties='Excel 8.0;HDR=YES'";
                }

                if (ext == ".xlsx")
                {
                    strCn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName
                            + ";Extended Properties='Excel 12.0 Xml;HDR=YES' ";
                }

                if (ext == ".xlsb")
                {
                    strCn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName
                            + ";Extended Properties='Excel 12.0;HDR=YES' ";
                }

                if (ext == ".xlsm")
                {
                    strCn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName
                            + ";Extended Properties='Excel 12.0 Macro;HDR=YES' ";
                }
            }

            using (var cn = new OleDbConnection(strCn))
            {
                var cmd = new OleDbCommand(queryCreateExcelTable, cn);
                cn.Open();
                cmd.ExecuteNonQuery();

                using (var da = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", cn))
                {
                    using (var cb = new OleDbCommandBuilder(da) { QuotePrefix = "[", QuoteSuffix = "]" })
                    {
                        // creates the INSERT INTO command
                        cmd = cb.GetInsertCommand();

                        // gets a hold of the INSERT INTO command.
                        foreach (DataRow row in selectedRows)
                        {
                            foreach (OleDbParameter param in cmd.Parameters)
                            {
                                param.Value = row[colNames[param.SourceColumn.Replace('#', '.')]];
                            }

                            cmd.ExecuteNonQuery(); // INSERT INTO command.
                        }
                    }
                }

                cmd.Dispose();
            }

            dataTable.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Tests the export2.
        /// </summary>
        /// <param name="table">
        /// The table.
        /// </param>
        /// <param name="sheetName">
        /// Name of the sheetName.
        /// </param>
        /// <param name="fileName">
        /// Name of the file.
        /// </param>
        public static void TestExport2(DataTable table, string sheetName, string fileName)
        {
            // Get an excel instance
            var excel = new Application();

            // Get a workbook
            Workbook wb = excel.Workbooks.Add();

            // Get a worksheet
            Worksheet ws = wb.Worksheets.Add();
            ws.Name = sheetName;

            // Add column names to the first row
            int col = 1;
            foreach (DataColumn c in table.Columns)
            {
                ws.Cells[1, col] = c.ColumnName;
                col++;
            }

            // Create a 2D array with the data from the table
            int i = 0;
            var data = new string[table.Rows.Count, table.Columns.Count];

            foreach (DataRow row in table.Rows)
            {
                var j = 0;
                foreach (DataColumn c in table.Columns)
                {
                    data[i, j] = row[c].ToString();
                    j++;
                }

                i++;
            }

            // Set the range value to the 2D array
            ws.Range[ws.Cells[2, 1], ws.Cells[table.Rows.Count + 1, table.Columns.Count]].Value2 = data;

            // Auto fit columns and rows, show excel, save.. etc
            excel.Columns.AutoFit();
            excel.Rows.AutoFit();
            excel.Visible = true;
        }

        #endregion
    }
}