#region

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

#endregion

namespace ExcelReportsDatastore
{
    /// <summary>
    /// The excel open xml reader.
    /// </summary>
    public class ExcelOpenXmlReader
    {
        #region Static Fields

        /// <summary>
        /// The shared string items.
        /// </summary>
        private static SharedStringItem[] sharedStringItems;

        /// <summary>
        /// The _table name.
        /// </summary>
        private static string tableName = string.Empty;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The get column index from name.
        /// </summary>
        /// <param name="columnName">
        /// The column name.
        /// </param>
        /// <returns>
        /// Returns the null or integer value.
        /// </returns>
        public static int? GetColumnIndexFromName(string columnName)
        {
            // return columnIndex;
            string name = columnName;

            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            int number = 0;
            int pow = 1;

            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">
        /// Address of the cell (example B2)
        /// </param>
        /// <returns>
        /// Column Name (example B)
        /// </returns>
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        /// <summary>
        /// Reads the excel file and store the data to the local database.
        /// </summary>
        /// <param name="fileName">
        /// Name of the file.
        /// </param>
        /// <returns>
        /// Returns database data using data table.
        /// </returns>
        public static DataTable ReadExcelFile(string fileName)
        {
            var localDatabaseConnection = ExcelOleDbReader.GetLocalConnection();
            localDatabaseConnection.Open();

            try
            {
                // Open the file. You can pass 'false', if you just need to open the file for reading.
                using (var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    var totalColumns = 0;

                    var stopWatch = new Stopwatch();
                    stopWatch.Start();

                    foreach (Sheet sheet in workbookPart.Workbook.Descendants<Sheet>())
                    {
                        var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;

                        sharedStringItems =
                            workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
                                .ToArray<SharedStringItem>();

                        var rowIndex = 0;

                        // Create a OpenXmlReader that can iterate through the worksheet parts and read the values in it.
                        var reader = OpenXmlReader.Create(worksheetPart);

                        while (reader.Read())
                        {
                            if (reader.ElementType != typeof(Row))
                            {
                                continue;
                            }

                            var row = (Row)reader.LoadCurrentElement();

                            var cells = row.Descendants<Cell>().ToArray();

                            if (rowIndex > 0)
                            {
                                StoreExcelRecordToDatabaseTable(localDatabaseConnection, cells, totalColumns);
                            }
                            else
                            {
                                totalColumns = CreateTableWithHeaderFromExcel(cells, localDatabaseConnection);
                            }

                            rowIndex++;
                        }

                        break;
                    }

                    stopWatch.Stop();

                    // Get the elapsed time as a TimeSpan value.
                    TimeSpan ts = stopWatch.Elapsed;

                    // Format and display the TimeSpan value.
                    string elapsedTime = string.Format(
                        "{0:00}:{1:00}:{2:00}.{3:00}", 
                        ts.Hours, 
                        ts.Minutes, 
                        ts.Seconds, 
                        ts.Milliseconds / 10);
                }

                var dissectQuery = string.Format("select top 10000 * from ExcelDataTable");

                return DatabaseReader.GetDatabaseDataTable(localDatabaseConnection, dissectQuery);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                localDatabaseConnection.Close();
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates the table with header from excel.
        /// </summary>
        /// <param name="cells">
        /// The cells.
        /// </param>
        /// <param name="localDatabaseConnection">
        /// The local database connection.
        /// </param>
        /// <returns>
        /// Returns the total number of columns created in a table.
        /// </returns>
        private static int CreateTableWithHeaderFromExcel(
            IEnumerable<Cell> cells, 
            SqlCeConnection localDatabaseConnection)
        {
            var columnNames =
                (from tempCell in cells where tempCell.CellValue != null select GetCellValue(tempCell)).Aggregate(
                    "ID", 
                    (current, cellValue) =>
                    string.IsNullOrEmpty(current) ? cellValue : string.Format("{0}|{1}", current, cellValue));

            /*foreach (var tempCell in cells)
            {
                if (tempCell.CellValue == null)
                {
                    continue;
                }

               var cellValue = GetCellValue(tempCell);

                columnNames = string.IsNullOrEmpty(columnNames)
                                    ? cellValue
                                    : string.Format("{0}|{1}", columnNames, cellValue);

              }*/
            var columnsArray = columnNames.Split('|');
            var totalColumns = columnsArray.Length;

            tableName = ExcelOleDbReader.CreateSqlTableFromExcelColumns(localDatabaseConnection, columnsArray);
            return totalColumns;
        }

        /*private string GetCellValue(Row row)
        {
            Cell theCell = row.Descendants<Cell>().Where(c => c.CellReference == ExcelColumnFromNumber(1) + row.RowIndex.ToString()).FirstOrDefault();

            String theCellValue = "";

            if (theCell != null)
            {
                theCellValue = theCell.InnerText;
            }
        }*/

        /*// Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        } */

        /// <summary>
        /// The get cell value.
        /// </summary>
        /// <param name="cell">
        /// The cell.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private static string GetCellValue(Cell cell)
        {
            if (cell.CellValue == null)
            {
                return null;
            }

            var value = cell.CellValue.InnerXml;

            /*if (cell.DataType == null || cell.DataType.Value != CellValues.SharedString)
            {
                return value;
            }*/
            if (cell.DataType == null)
            {
                return value;
            }

            SharedStringItem ssi = sharedStringItems[int.Parse(cell.CellValue.InnerText)];

            return ssi != null ? ssi.InnerText : null;

            // return sharedStringTable.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
        }

        /// <summary>
        /// Stores the excel record to database table.
        /// </summary>
        /// <param name="localDatabaseConnection">
        /// The local database connection.
        /// </param>
        /// <param name="cells">
        /// The cells.
        /// </param>
        /// <param name="totalColumns">
        /// Total Number of Columns
        /// </param>
        /// <exception cref="System.ArgumentNullException">
        /// cells
        /// </exception>
        private static void StoreExcelRecordToDatabaseTable(
            SqlCeConnection localDatabaseConnection, 
            IEnumerable<Cell> cells, 
            int totalColumns)
        {
            if (cells == null)
            {
                throw new ArgumentNullException("cells");
            }

            using (var cmd = new SqlCeCommand())
            {
                cmd.Connection = localDatabaseConnection;
                cmd.CommandText = tableName;
                cmd.CommandType = CommandType.TableDirect;

                using (
                    var databaseRecord = cmd.ExecuteResultSet(ResultSetOptions.Updatable | ResultSetOptions.Scrollable))
                {
                    SqlCeUpdatableRecord record = databaseRecord.CreateRecord();

                    var cellIndex = 1;

                    foreach (var tempCell in cells)
                    {
                        var columnIndexFromName = GetColumnIndexFromName(GetColumnName(tempCell.CellReference));

                        if (columnIndexFromName != null)
                        {
                            var cellColumnIndex = (int)columnIndexFromName;
                            cellColumnIndex--; // zero based index

                            while (cellIndex < cellColumnIndex)
                            {
                                /* if (cellIndex > totalColumns)
                                {
                                    // Ignore data outside the boundary of columns.
                                    break;
                                }*/

                                // Insert blank data here;
                                record.SetValue(cellIndex, null);
                                cellIndex++;
                            }
                        }

                        /*if (cellIndex > totalColumns)
                        {
                            // Ignore data outside the boundary of columns.
                            continue;
                        }*/
                        var cellValue = GetCellValue(tempCell);

                        record.SetValue(cellIndex, cellValue);
                        cellIndex++;
                    }

                    databaseRecord.Insert(record);
                }
            }
        }

        #endregion
    }
}