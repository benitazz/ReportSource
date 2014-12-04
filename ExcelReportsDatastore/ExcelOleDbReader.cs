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
using System.Reflection;
using System.Text;

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
    ///   The excel ole database reader.
    /// </summary>
    public class ExcelOleDbReader
    {
        #region Static Fields

        /// <summary>
        /// The sheets.
        /// </summary>
        private static List<string> sheets;

        /// <summary>
        /// The total columns.
        /// </summary>
        private static int totalColums;

        /// <summary>
        /// The _background worker.
        /// </summary>
        private static BackgroundWorker backgroundWorker;

        /// <summary>
        /// The book
        /// </summary>
        private static Workbook book;
        
        /// <summary>
        /// The sheet
        /// </summary>
        private static Worksheet sheet;

        /// <summary>
        /// The excel application.
        /// </summary>
        private static Application excelApplication;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Creates the SQL table from excel columns.
        /// </summary>
        /// <param name="sqlCeConnection">
        /// The SQL connection.
        /// </param>
        /// <param name="columns">
        /// The columns.
        /// </param>
        /// <returns>
        /// Return table name.
        /// </returns>
        public static string CreateSqlTableFromExcelColumns(SqlCeConnection sqlCeConnection, string[] columns)
        {
            const string TableName = "ExcelDataTable";

            DropTable(sqlCeConnection, TableName);

            // var columns = "[ExcelReportID] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY";
            var columnsNames = string.Empty;
            int totalColunms = 0;

            foreach (string column in columns)
            {
                var columnName = column;

                if (string.IsNullOrEmpty(columnName))
                {
                    continue;
                }

                columnName = columnName.ReplaceSpecialCharaterWithUnderScore();

                columnName = columnName.Replace(" ", "_");

                if (columnName.Equals("Group"))
                {
                    columnName = "Groups";
                }

                if (columnsNames == "F1" || columnsNames == "F2")
                {
                    continue;
                }

                if (columnName == "ID")
                {
                    columnsNames = string.IsNullOrEmpty(columnsNames)
                                       ? string.Format("[{0}] INT IDENTITY(1,1) PRIMARY KEY", columnName)
                                       : string.Format(
                                           "{0}, [{1}] INT IDENTITY(1,1) PRIMARY KEY", 
                                           columnsNames, 
                                           columnName);

                    ++totalColunms;

                    continue;
                }

                ++totalColunms;
                columnsNames = string.IsNullOrEmpty(columnsNames)
                                   ? string.Format("[{0}] NVARCHAR(500) NULL", columnName)
                                   : string.Format("{0}, [{1}] NVARCHAR(500)", columnsNames, columnName);
            }

            totalColums = totalColunms;
            var tableQuery = string.Format("CREATE TABLE {0}({1});", TableName, columnsNames);

            using (var command = new SqlCeCommand(tableQuery, sqlCeConnection))
            {
                command.ExecuteNonQuery();
            }

            return TableName;
        }

        /// <summary>
        /// Gets the excel data table.
        /// </summary>
        /// <param name="filename">
        /// The filename.
        /// </param>
        /// <param name="background">The background thread.</param>
        /// <returns>
        /// The <see cref="DataView"/>.
        /// </returns>
        public static DataTable GetExcelDataTable(string filename, BackgroundWorker background)
        {
            backgroundWorker = background;

            backgroundWorker.DoWork += BackgroundWorkerDoWork;

            background.RunWorkerAsync(filename);

            return null;

            // return LoadExcelDataToDatabase(filename);

            // SaveFileToDatabase(filename);
            /*var sheet1 = new DataTable();
            var csbuilder = new OleDbConnectionStringBuilder
                                {
                                    Provider = "Microsoft.ACE.OLEDB.12.0", 
                                    DataSource = filename
                                };
            csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=YES");*/

            /* var connect =
              "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= {0}; Extended Properties=\"Excel 12.0;IMEX=1;HDR=YES;TypeGuessRows=0;ImportMixedTypes=Text\"";*/
            /*var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename
                                   + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0\"";

            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                var defaultSheet = GetSheetName(connection);

                string selectSql = string.Format(@"SELECT * FROM [{0}]", defaultSheet);

                // MessageBox.Show(selectSql);
                using (var adapter = new OleDbDataAdapter(selectSql, connection))
                {
                    adapter.Fill(sheet1);

                    // dataGridView1.DataSource = sheet1;
                }

                connection.Close();
            }

            /*foreach (DataRow dr in sheet1.Rows)
            {
                foreach (DataColumn col in sheet1.Columns)
                {
                    if (col.ColumnNameFilter == "colName")
                    {
                        dr[col] = dr[col].ToString().Replace(" ", string.Empty);
                    }
                    else if (col.DataType == typeof(string))
                    {
                        dr[col] = dr[col].ToString().Trim();
                    }
                }
            }*/

            // return sheet1;
        }

        /// <summary>
        /// Handles the DoWork event of the backgroundWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        public static void BackgroundWorkerDoWork(object sender, DoWorkEventArgs e)
        {
            var dissectHelper = e.Argument as IReportData;

            if (dissectHelper != null)
            {
                backgroundWorker.ReportProgress(1);
                ExcelWriter.Dissect(dissectHelper);
                e.Result = dissectHelper;
                return;
            }

            var filename = e.Argument.ToString();
          // e.Result = LoadExcelDataToDatabase(filename);
        }

        /// <summary>
        /// Gets the local connection.
        /// </summary>
        /// <returns>
        /// Returns local database connection.
        /// </returns>
        /// <exception cref="System.Exception">Could not create a connection to the local database</exception>
        public static SqlCeConnection GetLocalConnection()
        {
            /*const string DatabaseConnectionString =
                @"Data Source=C:\Users\bbdnet1087\Documents\Visual Studio 2012\Projects\ExcelReportsGenerator\ExcelReportsDatastore\ExcelReportsDatabase.sdf;Max Database Size=4091;Max Buffer Size=4091;Persist Security Info=False;";*/

            try
            {
                return new SqlCeConnection(ConnectionStringHelper.ExcelDatabaseConnectionString);
            }
            catch (Exception exception)
            {
                throw new Exception("Could not create a connection to the local db");
            }
        }

        /// <summary>
        /// Gets the sheet names.
        /// </summary>
        /// <returns>
        /// returns list of sheet names.
        /// </returns>
        public static List<string> GetSheetNames()
        {
            return sheets;
        }

        /// <summary>
        /// Gets the sheet names.
        /// </summary>
        /// <param name="excelFile">
        /// The excel file.
        /// </param>
        /// <returns>
        /// Returns a list Sheet names.
        /// </returns>
        public static List<string> GetSheetNames(string excelFile)
        {
            var tableList = new List<string>();
            string excelColumns = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";"
                                  + "Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";

            using (var connection = new OleDbConnection(excelColumns))
            {
                connection.Open();
                DataTable tableInfo = connection.GetSchema("Tables");

                tableList.AddRange(from DataRow row in tableInfo.Rows select row["TABLE_NAME"].ToString());

                connection.Close();
            }

            return tableList;
        }

        /// <summary>
        /// Loads the excel data to database.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="background">The background.</param>
        /// <returns>
        /// Returns data table.
        /// </returns>
        /// <exception cref="System.Exception"></exception>
        public static DataTable LoadExcelDataToDatabase(string filePath, BackgroundWorker background)
        {
            backgroundWorker = background;
            backgroundWorker.ReportProgress(1);
            var stopWatch = new Stopwatch();
            stopWatch.Start();

            excelApplication = new Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };

            try
            {
                book = excelApplication.Workbooks.Open(
                    filePath, 
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
               
                // @"Data Source=C:\Users\bbdnet1087\Documents\Visual Studio 2012\Projects\ExcelReportsGenerator\ExcelReportsDatastore\ExcelReportsDatabase.sdf;Max Database Size=4091;Max Buffer Size=4091;Persist Security Info=False;";

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

                    StoreDataCsvToDatabase(sqlCeConnection, TempOutputFile, filePath);
                }

               stopWatch.Stop();

                // Get the elapsed time as a TimeSpan value.
                TimeSpan ts = stopWatch.Elapsed;

                // Format and display the TimeSpan value.
                string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", 
                    ts.Hours, 
                    ts.Minutes, 
                    ts.Seconds, 
                    ts.Milliseconds / 10);

                var dissectQuery = string.Format("select top 10000 * from ExcelDataTable");

               /* string DatabaseConnectionString2 = ConnectionStringHelper.GetConnectionString();*/

                    // @"Data Source=C:\Users\bbdnet1087\Documents\Visual Studio 2012\Projects\ExcelReportsGenerator\ExcelReportsDatastore\ExcelReportsDatabase.sdf;Max Database Size=4091;Max Buffer Size=4091;Persist Security Info=False;";

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
                //range = null;
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

        /// <summary>
        /// Saves the file to database.
        /// </summary>
        /// <param name="filePath">
        /// The file path.
        /// </param>
        public static void SaveFileToDatabase(string filePath)
        {
           var excelConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath
                                  + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0\"";

            // string connectionString = @"Data Source=(LocalDb)\v11.0;Initial Catalog=ExcelReportsDb;Integrated Security=True; MultipleActiveResultSets=True";
            /*const string DatabaseConnectionString =
                @"Data Source=C:\Users\bbdnet1087\Documents\Visual Studio 2012\Projects\ExcelReportsGenerator\ExcelReportsDatastore\ExcelReportsDatabase.sdf;Max Database Size=4091;Max Buffer Size=4091;Persist Security Info=False;";*/

            try
            {
                using (var sqlCeConnection = new SqlCeConnection(ConnectionStringHelper.ExcelDatabaseConnectionString))
                {
                    sqlCeConnection.Open();

                    using (var excelConnection = new OleDbConnection(excelConnString))
                    {
                        excelConnection.Open();
                        var defaultSheet = GetSheetName(excelConnection);

                        var tableName = CreateSqlTableFromExcelColumns(excelConnection, sqlCeConnection);

                        var selectSql = string.Format(@"SELECT * FROM [{0}]", defaultSheet);

                        // Create OleDbCommand to fetch data from Excel 
                        using (var cmd = new OleDbCommand(selectSql, excelConnection))
                        {
                            using (OleDbDataReader oleDbDataReader = cmd.ExecuteReader())
                            {
                                StoreData(sqlCeConnection, tableName, oleDbDataReader);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates the SQL table from excel columns.
        /// </summary>
        /// <param name="excelConnection">
        /// The excel connection.
        /// </param>
        /// <param name="sqlCeConnection">
        /// The SQL connection.
        /// </param>
        /// <returns>
        /// Return table name.
        /// </returns>
        private static string CreateSqlTableFromExcelColumns(
            OleDbConnection excelConnection, 
            SqlCeConnection sqlCeConnection)
        {
            DataTable columnsInfo = excelConnection.GetSchema("Columns");

            const string TableName = "ExcelDataTable";

            DropTable(sqlCeConnection, TableName);

            // var columns = "[ExcelReportID] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY";
            var columns = string.Empty;

            int totalColunms = 0;

            for (var i = 0; i < columnsInfo.Rows.Count; i++)
            {
                var columnName = columnsInfo.Rows[i].ItemArray[3].ToString();

                if (string.IsNullOrEmpty(columnName))
                {
                    continue;
                }

                columnName = columnName.Replace(" ", "_");

                if (columnName.Equals("Group"))
                {
                    columnName = "Groups";
                }

                if (columns.Contains(columnName) || columnName == "F1" || columnName == "F2")
                {
                    continue;
                }

                ++totalColunms;
                columns = string.IsNullOrEmpty(columns)
                              ? string.Format("[{0}] NVARCHAR(250) NULL", columnName)
                              : string.Format("{0}, {1} NVARCHAR(250)", columns, columnName);
            }

            totalColums = totalColunms;
            var tableQuery = string.Format("CREATE TABLE {0}({1});", TableName, columns);

            using (var command = new SqlCeCommand(tableQuery, sqlCeConnection))
            {
                command.ExecuteNonQuery();
            }

            return TableName;
        }

        /// <summary>
        /// Drops the table.
        /// </summary>
        /// <param name="sqlCeConnection">
        /// The SQL connection.
        /// </param>
        /// <param name="tableName">
        /// Name of the table.
        /// </param>
        public static void DropTable(SqlCeConnection sqlCeConnection, string tableName)
        {
            var dt = sqlCeConnection.GetSchema("tables");

            var tableExists = dt.Rows.Cast<DataRow>().Any(row => row["TABLE_NAME"].ToString() == tableName);

            if (!tableExists)
            {
                return;
            }

            using (var cmd = new SqlCeCommand(string.Format("DROP TABLE {0}", tableName), sqlCeConnection))
            {
                cmd.ExecuteNonQuery();
            }
        }

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
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="record">The record.</param>
        /// <param name="totalColumns">The total columns.</param>
        private static void GetExcelRowValuesByIndex(
            Worksheet worksheet,
            int rowIndex,
            SqlCeUpdatableRecord record,
            int totalColumns)
        {
            try
            {
                var startCellRange = string.Format("A{0}", rowIndex);
                var endCellRange = string.Format("{0}{1}", GetExcelColumnName(totalColumns), rowIndex);

                Range range = worksheet.Range[startCellRange, endCellRange];

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
        /// Gets the name of the sheet.
        /// </summary>
        /// <param name="excelConnection">
        /// The excel connection.
        /// </param>
        /// <returns>
        /// Returns the sheet name from the selected worksheet.
        /// </returns>
        private static string GetSheetName(OleDbConnection excelConnection)
        {
            DataTable tableInfo = excelConnection.GetSchema("Tables");

            var tableList = new List<string>();

            tableList.AddRange(
                from DataRow row in tableInfo.Rows
                where row["TABLE_NAME"] != null && !row["TABLE_NAME"].ToString().Contains("_xlnm#_FilterDatabase")
                select row["TABLE_NAME"].ToString());

            sheets = tableList;
            var defaultSheet = tableList[0];

            return defaultSheet;
        }

        /// <summary>
        /// Stores the data.
        /// </summary>
        /// <param name="sqlCeConnection">
        /// The SQL connection.
        /// </param>
        /// <param name="tableName">
        /// Name of the table.
        /// </param>
        /// <param name="dataReader">
        /// Data Reader
        /// </param>
        private static void StoreData(SqlCeConnection sqlCeConnection, string tableName, OleDbDataReader dataReader)
        {
            if (sqlCeConnection.State == ConnectionState.Closed)
            {
                sqlCeConnection.Open();
            }

            using (var cmd = new SqlCeCommand())
            {
                cmd.Connection = sqlCeConnection;
                cmd.CommandText = tableName;
                cmd.CommandType = CommandType.TableDirect;

                using (var rs = cmd.ExecuteResultSet(ResultSetOptions.Updatable | ResultSetOptions.Scrollable))
                {
                    SqlCeUpdatableRecord record = rs.CreateRecord();

                    while (dataReader.Read())
                    {
                        for (int index = 0; index < totalColums; index++)
                        {
                            record.SetValue(index, dataReader[index]);
                        }

                        rs.Insert(record);
                    }
                }
            }
        }

        /// <summary>
        /// Stores the data CSV file to database.
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
        private static void StoreDataCsvToDatabase(SqlCeConnection sqlCeConnection, string filePath, string excelPath)
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

                    tableName = CreateSqlTableFromExcelColumns(sqlCeConnection, splitResult);

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

        /// <summary>
        /// The deleteme.
        /// </summary>
        /// <param name="excelFilePath">
        /// The excel file path.
        /// </param>
        /// <exception cref="Exception">
        /// </exception>
        private static void deleteme(string excelFilePath)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            excelApplication = new Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };
            
            Range range = null;

            // the range object is used to hold the data
            // we'll be reading from and to find the range of data.
            try
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

                sheet = (Worksheet)book.Worksheets[1];

                range = sheet.Range["A1", Missing.Value];

                /*range = range.End[XlDirection.xlToRight];

                range = range.End[XlDirection.xlDown];*/
                range = range.SpecialCells(XlCellType.xlCellTypeLastCell);

                string mainDownAddress = range.Address[false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing];

                range = sheet.Range["A1", mainDownAddress];

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    var startCellRange = string.Format("A{0}", rowIndex);

                    range = sheet.Range[startCellRange, Missing.Value];

                    range = range.End[XlDirection.xlToRight];

                    string downAddress = range.Address[false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing];

                    range = sheet.Range[startCellRange, downAddress];

                    var values = (object[,])range.Value2;

                    columnCount = range.Columns.Count;

                    for (int i = 1; i <= columnCount; i++)
                    {
                        var val = values[1, i];

                        var cellValues = val != null ? val.ToString() : null;
                    }
                }

                stopWatch.Stop();

                // Get the elapsed time as a TimeSpan value.
                TimeSpan ts = stopWatch.Elapsed;

                // Format and display the TimeSpan value.
                string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", 
                    ts.Hours, 
                    ts.Minutes, 
                    ts.Seconds, 
                    ts.Milliseconds / 10);
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
            finally
            {
                range = null;
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
    }
}