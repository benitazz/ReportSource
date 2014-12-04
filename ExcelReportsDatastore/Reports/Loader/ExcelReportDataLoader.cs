#region

using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

using ExcelReportsDatastore.Helpers;
using ExcelReportsDatastore.Reports.Data;

using Microsoft.Office.Interop.Excel;

using DataTable = System.Data.DataTable;

#endregion

namespace ExcelReportsDatastore.Reports.Loader
{
    /// <summary>
    /// The load report data.
    /// </summary>
    public class ExcelReportDataLoader : IReportLoader
    {
        #region Static Fields

        /// <summary>
        /// The _background worker.
        /// </summary>
        private static BackgroundWorker backgroundWorker;

        /// <summary>
        /// The book
        /// </summary>
        private static Workbook book;

        /// <summary>
        /// The excel application.
        /// </summary>
        private static Application excelApplication;

        /// <summary>
        /// The sheet
        /// </summary>
        private static Worksheet sheet;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Loads the report data.
        /// </summary>
        /// <param name="reportData">
        /// The report data.
        /// </param>
        /// <param name="worker">
        /// The worker.
        /// </param>
        /// <returns>
        /// The <see cref="DataTable"/>.
        /// </returns>
        public DataTable LoadReportData(IReportData reportData, BackgroundWorker worker)
        {
            backgroundWorker = worker;
            backgroundWorker.ReportProgress(1);
            excelApplication = new Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };

            try
            {
                book = excelApplication.Workbooks.Open(
                    reportData.FileName, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value, 
                    Missing.Value);

                const string TempOutputFile = @"C:\Temp\output.csv";

                if (File.Exists(TempOutputFile))
                {
                    File.Delete(TempOutputFile);
                }

                book.SaveAs(TempOutputFile, XlFileFormat.xlCurrentPlatformText);
                book.Close(false);
                excelApplication.Quit();
                book = null;
                excelApplication = null;

                using (var sqlCeConnection = new SqlCeConnection(ConnectionStringHelper.ExcelDatabaseConnectionString))
                {
                    try
                    {
                        if (sqlCeConnection.State == ConnectionState.Closed)
                        {
                            sqlCeConnection.Open();
                        }
                    }
                    catch (Exception exception)
                    {
                        // engine.Repair(null, RepairOption.DeleteCorruptedRows);
                        Console.WriteLine(exception);
                    }

                    StoreCsvDataToDatabase(sqlCeConnection, TempOutputFile, reportData.FileName);
                }

                var dissectQuery = string.Format("select top 10000 * from ExcelDataTable");

                using (var sqlCeConnection = new SqlCeConnection(ConnectionStringHelper.ExcelDatabaseConnectionString))
                {
                    sqlCeConnection.Open();

                    return DatabaseReader.GetDatabaseDataTable(sqlCeConnection, dissectQuery);
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
            finally
            {
                sheet = null;
                if (book != null)
                {
                    book.Close(false, Missing.Value, Missing.Value);
                }

                book = null;

                if (excelApplication != null)
                {
                    excelApplication.Quit();
                }

                excelApplication = null;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Gets the name of the excel column.
        /// </summary>
        /// <param name="columnNumber">
        /// The column number.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// Gets the index of the excel row values by.
        /// </summary>
        /// <param name="excelFilePath">
        /// The excel file path.
        /// </param>
        /// <param name="rowIndex">
        /// Index of the row.
        /// </param>
        /// <param name="record">
        /// The record.
        /// </param>
        /// <param name="totalColumns">
        /// total number of columns.
        /// </param>
        /// <exception cref="System.Exception">
        /// Throws the exception message.
        /// </exception>
        private static void GetExcelRowValuesByIndex(
            string excelFilePath, 
            int rowIndex, 
            SqlCeUpdatableRecord record, 
            int totalColumns)
        {
            if (excelApplication == null)
            {
                excelApplication = new Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };
            }

            try
            {
                if (book == null)
                {
                    book = excelApplication.Workbooks.Open(
                        excelFilePath, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value, 
                        Missing.Value);
                }

                if (sheet == null)
                {
                    sheet = (Worksheet)book.Worksheets[1];
                }

                var startCellRange = string.Format("A{0}", rowIndex);
                var endCellRange = string.Format("{0}{1}", GetExcelColumnName(totalColumns), rowIndex);

                Range range = sheet.Range[startCellRange, endCellRange];

                int columnCount = range.Columns.Count;

                var values = (object[,])range.Value2;

                var recordIndex = 0;

                for (int i = 1; i <= columnCount; i++)
                {
                    var val = values[1, i];

                    record.SetValue(recordIndex, val != null ? val.ToString() : null);

                    recordIndex++;
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
        }

        /// <summary>
        /// Stores the CSV data file to database.
        /// </summary>
        /// <param name="sqlCeConnection">
        /// The SQL connection.
        /// </param>
        /// <param name="filePath">
        /// The file path.
        /// </param>
        /// <param name="excelPath">
        /// The excel Path.
        /// </param>
        private static void StoreCsvDataToDatabase(SqlCeConnection sqlCeConnection, string filePath, string excelPath)
        {
            if (sqlCeConnection.State == ConnectionState.Closed)
            {
                sqlCeConnection.Open();
            }

            string tableName = string.Empty;
            var tolalColumns = 0;

            using (var reader = new StreamReader(filePath, Encoding.Default))
            {
                while (!reader.EndOfStream)
                {
                    string message = reader.ReadLine();

                    if (message == null)
                    {
                        continue;
                    }

                    string[] splitResult = message.Split(new[] { '\t' }, StringSplitOptions.None);

                    tableName = ExcelOleDbReader.CreateSqlTableFromExcelColumns(sqlCeConnection, splitResult);

                    tolalColumns = splitResult.Length;
                    break;
                }
            }

            using (var cmd = new SqlCeCommand())
            {
                cmd.Connection = sqlCeConnection;
                cmd.CommandText = tableName;
                cmd.CommandType = CommandType.TableDirect;

                using (var rs = cmd.ExecuteResultSet(ResultSetOptions.Updatable | ResultSetOptions.Scrollable))
                {
                    SqlCeUpdatableRecord record = rs.CreateRecord();

                    using (var reader = new StreamReader(filePath, Encoding.Default))
                    {
                        var rowCount = 0;

                        while (!reader.EndOfStream)
                        {
                            string message = reader.ReadLine();

                            /*var percentComplete = (int)(rowCount / (float)tolalColumns * 100);
                            backgroundWorker.ReportProgress(percentComplete);*/
                            if (message == null || rowCount == 0)
                            {
                                ++rowCount;
                                continue;
                            }

                            string[] splitResult = message.Split(new[] { '\t' }, StringSplitOptions.None);

                            // Read One Row and Sep
                            var templist = splitResult.ToList();

                            var list =
                                (from temp in templist where temp != "\"" select temp.Replace("\"", string.Empty))
                                    .ToList();

                            if (list.Count > tolalColumns)
                            {
                                var extraColumnCount = tolalColumns + 1;

                                if (list.Count != extraColumnCount && !string.IsNullOrEmpty(list[tolalColumns]))
                                {
                                    ++rowCount;

                                    GetExcelRowValuesByIndex(excelPath, rowCount, record, tolalColumns);
                                    rs.Insert(record);
                                    continue;
                                }
                            }

                            // record.SetValues(list.ToArray());
                            for (int index = 0; index < tolalColumns; index++)
                            {
                                record.SetValue(index, list[index]);
                            }

                            rs.Insert(record);

                            ++rowCount;
                        }
                    }
                }
            }
        }

        #endregion
    }
}