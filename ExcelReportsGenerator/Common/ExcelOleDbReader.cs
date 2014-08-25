#region

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using DataTable = System.Data.DataTable;

#endregion

namespace ExcelReportsGenerator.Common
{
  /// <summary>
  ///   The excel ole db reader.
  /// </summary>
  public class ExcelOleDbReader
  {
    /// <summary>
    /// The sheets.
    /// </summary>
    private static List<string> sheets;

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
      var sheet1 = new DataTable();
      var csbuilder = new OleDbConnectionStringBuilder();
      csbuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
      csbuilder.DataSource = filename;
      csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=YES");

      /*var connect =
        "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= {0}; Extended Properties=\"Excel 12.0;IMEX=1;HDR=YES;TypeGuessRows=0;ImportMixedTypes=Text\"";*/
      var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename
                             + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0\"";

      using (var connection = new OleDbConnection(connectionString))
      {
        connection.Open();
        DataTable tableInfo = connection.GetSchema("Tables");

        var tableList = new List<string>();

        tableList.AddRange(from DataRow row in tableInfo.Rows
                           where row["TABLE_NAME"] != null && 
                                !row["TABLE_NAME"].ToString().Contains("_xlnm#_FilterDatabase")
                           select row["TABLE_NAME"].ToString());

        sheets = tableList;

        var defaultSheet = tableList[0];

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

    #endregion
  }
}