#region

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlServerCe;
using System.IO;
using System.Linq;
using System.Windows;

using DataTable = System.Data.DataTable;

#endregion

namespace ExcelReportsGenerator.Common
{
  /// <summary>
  ///   The excel ole database reader.
  /// </summary>
  public class ExcelOleDbReader
  {
    /// <summary>
    /// The sheets.
    /// </summary>
    private static List<string> sheets;

    /// <summary>
    /// The total columns.
    /// </summary>
    private static int totalColums;

    #region Public Methods and Operators

    /// <summary>
    /// Gets the excel data table.
    /// </summary>
    /// <param name="filename">
    /// The filename.
    /// </param>
    /// <returns>
    /// The <see cref="DataView"/>.
    /// </returns>
    public static DataTable GetExcelDataTable(string filename)
    {
      SaveFileToDatabase(filename);

      var sheet1 = new DataTable();
      var csbuilder = new OleDbConnectionStringBuilder { Provider = "Microsoft.ACE.OLEDB.12.0", DataSource = filename };
      csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=YES");

      /*var connect =
        "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= {0}; Extended Properties=\"Excel 12.0;IMEX=1;HDR=YES;TypeGuessRows=0;ImportMixedTypes=Text\"";*/
      var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename
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

      foreach (DataRow dr in sheet1.Rows)
      {
        foreach (DataColumn col in sheet1.Columns)
        {
          if (col.ColumnName == "colName")
          {
            dr[col] = dr[col].ToString().Replace(" ", string.Empty);
          }
          else if (col.DataType == typeof(string))
          {
            dr[col] = dr[col].ToString().Trim();
          }
        }
      }

      return sheet1;
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
    /// Exports to XLSX.
    /// </summary>
    /// <param name="sheetToCreate">
    /// The sheet to create.
    /// </param>
    /// <param name="dtToExport">
    /// The dt to export.
    /// </param>
    /// <param name="tableName">
    /// Name of the table.
    /// </param>
    public static void ExportToXlsx(string sheetToCreate, DataTable dtToExport, string tableName)
    {
      var rows = new List<DataRow>();

      foreach (DataRow row in dtToExport.Rows)
      {
        rows.Add(row);
      }

      ExportToXlsx(sheetToCreate, rows, dtToExport, tableName);
    }

    /// <summary>
    /// Exports to XLSX.
    /// </summary>
    /// <param name="sheetToCreate">
    /// The sheet to create.
    /// </param>
    /// <param name="selectedRows">
    /// The selected rows.
    /// </param>
    /// <param name="dataTable">
    /// The data table.
    /// </param>
    /// <param name="tableName">
    /// Name of the table.
    /// </param>
    public static void ExportToXlsx(
      string sheetToCreate,
      List<DataRow> selectedRows,
      DataTable dataTable,
      string tableName)
    {
      const char Space = ' ';
      string dest = sheetToCreate;

      if (File.Exists(dest))
      {
        File.Delete(dest);
      }

      sheetToCreate = dest;

      if (tableName == null)
      {
        tableName = string.Empty;
      }

      tableName = tableName.Trim().Replace(Space, '_');

      if (tableName.Length == 0)
      {
        tableName = dataTable.TableName.Replace(Space, '_');
      }

      if (tableName.Length == 0)
      {
        tableName = "NoTableName";
      }

      if (tableName.Length > 30)
      {
        tableName = tableName.Substring(0, 30);
      }

      // Excel names are less than 31 chars
      string queryCreateExcelTable = "CREATE TABLE [" + tableName + "] (";
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
      if (selectedRows.Count > 65000 && sheetToCreate.ToLower().EndsWith(".xls"))
      {
        // use Excel 2007 for large sheets.
        sheetToCreate = sheetToCreate.ToLower().Replace(".xls", string.Empty) + ".xlsx";
      }

      string strCn = string.Empty;
      var extension = Path.GetExtension(sheetToCreate);

      if (extension != null)
      {
        string ext = extension.ToLower();

        if (ext == ".xls")
        {
          strCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sheetToCreate
                  + "; Extended Properties='Excel 8.0;HDR=YES'";
        }

        if (ext == ".xlsx")
        {
          strCn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sheetToCreate
                  + ";Extended Properties='Excel 12.0 Xml;HDR=YES' ";
        }

        if (ext == ".xlsb")
        {
          strCn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sheetToCreate
                  + ";Extended Properties='Excel 12.0;HDR=YES' ";
        }

        if (ext == ".xlsm")
        {
          strCn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sheetToCreate
                  + ";Extended Properties='Excel 12.0 Macro;HDR=YES' ";
        }
      }

      var cn = new OleDbConnection(strCn);
      var cmd = new OleDbCommand(queryCreateExcelTable, cn);
      cn.Open();
      cmd.ExecuteNonQuery();

      var da = new OleDbDataAdapter("SELECT * FROM [" + tableName + "]", cn);
      var cb = new OleDbCommandBuilder(da);

      // creates the INSERT INTO command
      cb.QuotePrefix = "[";
      cb.QuoteSuffix = "]";
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

      cn.Close();
      cn.Dispose();
      da.Dispose();
      GC.Collect();
      GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// Saves the file to database.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    public static void SaveFileToDatabase(string filePath)
    {
      var excelConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath
                           + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0\"";

     // string connectionString = @"Data Source=(LocalDb)\v11.0;Initial Catalog=ExcelReportsDb;Integrated Security=True; MultipleActiveResultSets=True";

      const string DatabaseConnectionString = @"Data Source=C:\Users\bbdnet1087\Documents\Visual Studio 2012\Projects\ExcelReportsGenerator\ExcelReportsDatastore\ExcelReportsDatabase.sdf;Max Database Size=4091;Max Buffer Size=4091;Persist Security Info=False;";

      try
      {
        using (var sqlCeConnection = new SqlCeConnection(DatabaseConnectionString))
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
        MessageBox.Show(exception.Message);
      }
    }

    /// <summary>
    /// Creates the SQL table from excel columns.
    /// </summary>
    /// <param name="excelConnection">The excel connection.</param>
    /// <param name="sqlCeConnection">The SQL connection.</param>
    /// <returns>
    /// Return table name.
    /// </returns>
    private static string CreateSqlTableFromExcelColumns(OleDbConnection excelConnection, SqlCeConnection sqlCeConnection)
    {
      DataTable columnsInfo = excelConnection.GetSchema("Columns");

      const string TableName = "ExcelDataTable";

      DropTable(sqlCeConnection, TableName);

      // var columns = "[ID] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY";

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
        columns = string.IsNullOrEmpty(columns) ?
          string.Format("[{0}] NVARCHAR(250) NULL", columnName) :
          string.Format("{0}, {1} NVARCHAR(250)", columns, columnName);
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
    /// <param name="sqlCeConnection">The SQL connection.</param>
    /// <param name="tableName">Name of the table.</param>
    private static void DropTable(SqlCeConnection sqlCeConnection, string tableName)
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
    /// Gets the name of the sheet.
    /// </summary>
    /// <param name="excelConnection">The excel connection.</param>
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
    /// <param name="sqlCeConnection">The SQL connection.</param>
    /// <param name="tableName">Name of the table.</param>
    /// <param name="dataReader">Data Reader</param>
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

    #endregion
  }
}