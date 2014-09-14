#region

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

using BenMVVM;

using ExcelReportsGenerator.Common;
using ExcelReportsGenerator.Common.Helpers;

using Microsoft.Win32;

#endregion

namespace ExcelReportsGenerator.ViewModels
{
  /// <summary>
  ///   The excel report view model.
  /// </summary>
  public class ExcelReportViewModel : ViewModelBase, IShellViewModel
  {
    #region Fields

    /// <summary>
    ///   The _columns names list
    /// </summary>
    private ObservableCollection<string> _columnsNamesList;

    /// <summary>
    ///   The _excel data.
    /// </summary>
    private DataView _excelData;

    /// <summary>
    ///   The _data table.
    /// </summary>
    private DataTable _excelSheetDataTable;

    /// <summary>
    ///   The _file name.
    /// </summary>
    private string _fileName;

    /// <summary>
    ///   The _is busy.
    /// </summary>
    private bool _isBusy;

    /// <summary>
    ///   The _is column filter enabled
    /// </summary>
    private bool _isColumnFilterEnabled;

    /// <summary>
    ///   The _selected column filter
    /// </summary>
    private string _selectedColumnFilter;

    /// <summary>
    ///   The _selected excel report.
    /// </summary>
    private ExcelReport _selectedExcelReport;

    /// <summary>
    ///   The _selected sheet
    /// </summary>
    private string _selectedSheet;

    /// <summary>
    ///   The _sheets
    /// </summary>
    private List<string> _sheets;

    /// <summary>
    ///   The _total records
    /// </summary>
    private int _totalRecords;

    #endregion

    #region Constructors and Destructors

    /// <summary>
    ///   Initializes a new instance of the <see cref="ExcelReportViewModel" /> class.
    /// </summary>
    public ExcelReportViewModel()
    {
      this.IsBusy = false;
    }

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets the browse file command.
    /// </summary>
    /// <value>
    ///   The browse file command.
    /// </value>
    public ICommand BrowseFileCommand
    {
      get
      {
        return new RelayCommand(param => this.FileBrowserHandler());
      }
    }

    /// <summary>
    ///   Gets or sets the columns names list.
    /// </summary>
    /// <value>
    ///   The columns names list.
    /// </value>
    public ObservableCollection<string> ColumnsNamesList
    {
      get
      {
        return this._columnsNamesList;
      }

      set
      {
        this._columnsNamesList = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the excel data.
    /// </summary>
    /// <value>
    ///   The excel data.
    /// </value>
    public DataView ExcelData
    {
      get
      {
        return this._excelData;
      }

      set
      {
        this._excelData = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets the quantity report command.
    /// </summary>
    /// <value>
    ///   The quantity report command.
    /// </value>
    public ICommand ExcelReportGeneratorCommand
    {
      get
      {
        return new RelayCommand(param => this.ExcelReportGenerator());
      }
    }

    /// <summary>
    ///   Gets or sets the name of the file.
    /// </summary>
    /// <value>
    ///   The name of the file.
    /// </value>
    public string FileName
    {
      get
      {
        return this._fileName;
      }

      set
      {
        this._fileName = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets a value indicating whether is busy.
    /// </summary>
    public bool IsBusy
    {
      get
      {
        return this._isBusy;
      }

      set
      {
        this._isBusy = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets a value indicating whether [is column filter enabled].
    /// </summary>
    /// <value>
    ///   <c>true</c> if [is column filter enabled]; otherwise, <c>false</c>.
    /// </value>
    public bool IsColumnFilterEnabled
    {
      get
      {
        return this._isColumnFilterEnabled;
      }

      set
      {
        this._isColumnFilterEnabled = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the selected column.
    /// </summary>
    /// <value>
    ///   The selected column.
    /// </value>
    public string SelectedColumnFilter
    {
      get
      {
        return this._selectedColumnFilter;
      }

      set
      {
        this._selectedColumnFilter = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the selected excel report.
    /// </summary>
    /// <value>
    ///   The selected excel report.
    /// </value>
    public ExcelReport SelectedExcelReport
    {
      get
      {
        return this._selectedExcelReport;
      }

      set
      {
        this._selectedExcelReport = value;

        if (this._selectedExcelReport == ExcelReport.ReplicateReport)
        {
          this.SelectedColumnFilter = "Quantity";
          this.IsColumnFilterEnabled = false;
        }
        else
        {
          this.SelectedColumnFilter = string.Empty;
          this.IsColumnFilterEnabled = true;
        }

        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the selected sheet.
    /// </summary>
    /// <value>
    ///   The selected sheet.
    /// </value>
    public string SelectedSheet
    {
      get
      {
        return this._selectedSheet;
      }

      set
      {
        this._selectedSheet = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the sheets.
    /// </summary>
    /// <value>
    ///   The sheets.
    /// </value>
    public List<string> Sheets
    {
      get
      {
        return this._sheets;
      }

      set
      {
        this._sheets = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the total records.
    /// </summary>
    /// <value>
    ///   The total records.
    /// </value>
    public int TotalRecords
    {
      get
      {
        return this._totalRecords;
      }

      set
      {
        this._totalRecords = value;
        this.NotifyPropertyChanged();
      }
    }

    #endregion

    #region Public Methods and Operators

    /// <summary>
    ///   Quantities the report.
    /// </summary>
    public void ExcelReportGenerator()
    {
      if (this._excelSheetDataTable == null)
      {
        return;
      }

      this.IsBusy = true;
      DataTable results = this._excelSheetDataTable.Copy();
      results.Rows.Clear();

      switch (this.SelectedExcelReport)
      {
        case ExcelReport.ReplicateReport:
          {
            this.ExcelReplicateReportGenerator(results);
            break;
          }

        case ExcelReport.DissectReport:
          {
            this.ExcelDissectReportGenerator(results);
            break;
          }

        default:
          {
            break;
          }
      }

      this.IsBusy = false;
    }

    /// <summary>
    ///   Files the browser handler.
    /// </summary>
    public void FileBrowserHandler()
    {
      try
      {
        this.IsBusy = true;

        var fileDialog = new OpenFileDialog
                           {
                             Filter =
                               "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx;*.txt|All Files(*.*)|*.*"
                           };

        this.FileName = FileBrowserHelper.GetfileName(fileDialog);

        if (string.IsNullOrEmpty(this.FileName))
        {
          this.IsBusy = false;
          return;
        }

        var fileExtension = Path.GetExtension(this.FileName);

        if (fileExtension.Contains(".xl"))
        {
          this._excelSheetDataTable = ExcelOleDbReader.GetExcelDataTable(this.FileName);

          this.ExcelData = this._excelSheetDataTable.DefaultView;

          this.TotalRecords = this.ExcelData.Count;

          this.ColumnsNamesList = new ObservableCollection<string>();

          foreach (DataColumn column in this._excelSheetDataTable.Columns)
          {
            this.ColumnsNamesList.Add(column.ColumnName);
          }

          this.NotifyPropertyChanged(() => this.ColumnsNamesList);
          this.IsColumnFilterEnabled = false;

          this.Sheets = ExcelOleDbReader.GetSheetNames();

          this.SelectedSheet = this.Sheets[0];
          this.SelectedColumnFilter = "Quantity";
        }
      }
      catch (Exception exception)
      {
        MessageBox.Show(exception.ToString());
      }

      this.IsBusy = false;
    }

    #endregion

    #region Methods

    /// <summary>
    /// Adds the row to data table.
    /// </summary>
    /// <param name="results">
    /// The results.
    /// </param>
    /// <param name="row">
    /// The row.
    /// </param>
    private static void AddRowToDatatable(DataTable results, DataRow row)
    {
      var newRow = results.NewRow();
      newRow.ItemArray = row.ItemArray;
      results.Rows.Add(newRow);
    }

    /// <summary>
    /// Excels the dissect report generator.
    /// </summary>
    /// <param name="results">
    /// The results.
    /// </param>
    private void ExcelDissectReportGenerator(DataTable results)
    {
      var view = new DataView(this._excelSheetDataTable);
      DataTable distinctValues = view.ToTable(true, this.SelectedColumnFilter);

      var directoryName = string.Empty;

      foreach (DataRow row in distinctValues.Rows)
      {
        var value = row[this.SelectedColumnFilter];

        var expression = string.Format("[{0}] = '{1}'", this.SelectedColumnFilter, value);

        DataRow[] filteredRows = this._excelSheetDataTable.Select(expression);

        foreach (var filteredRow in filteredRows)
        {
          AddRowToDatatable(results, filteredRow);
        }

        var dissectColumnName = value.ToString().Trim();

        dissectColumnName = this.RemoveSpecialCharacters(dissectColumnName);

         var fileName = Path.GetFileNameWithoutExtension(this.FileName);

          fileName = this.RemoveSpecialCharacters(fileName);

        directoryName = string.Format(@"C:\Dissect\{0}", fileName);
       FileBrowserHelper.CreateDirectory(directoryName);

        var filename = string.Format(@"{0}\{1}_{2}", directoryName, dissectColumnName, Path.GetFileName(this.FileName));
        
        ExcelOleDbReader.ExportToXlsx(filename, results, value.ToString());

        results.Rows.Clear();
      }

      // opens the folder in explorer
      Process.Start(directoryName);

      // opens the folder in explorer
      // Process.Start("explorer.exe", @"c:\temp");
    }

    /// <summary>
    /// Removes the special characters.
    /// </summary>
    /// <param name="str">The string.</param>
    /// <returns></returns>
    private string RemoveSpecialCharacters(string str)
    {
        return Regex.Replace(str, "[^a-zA-Z0-9_.]+", "_", RegexOptions.Compiled);

        /*var sb = new StringBuilder();

        foreach (char c in str)
        {
            if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
            {
                sb.Append(c);
            }
        }

        return sb.ToString();*/
    }

    /// <summary>
    /// Excels the replicate report generator.
    /// </summary>
    /// <param name="results">
    /// The results.
    /// </param>
    private void ExcelReplicateReportGenerator(DataTable results)
    {
      var index = 0;

      foreach (DataRow row in this._excelSheetDataTable.Rows)
      {
        var value = this._excelSheetDataTable.Rows[index++][this.SelectedColumnFilter];

        if (value == null)
        {
          continue;
        }

        var quantity = int.Parse(value.ToString());

        if (quantity == 0 || quantity == 1)
        {
          AddRowToDatatable(results, row);

          continue;
        }

        for (int i = 0; i < quantity; i++)
        {
          AddRowToDatatable(results, row);
        }
      }

      // var fileName = string.Format("{0}_{1}", DateTime.Now, Path.GetFileName(this.FileName));
      var fileName = @"c:\test.xlsx";

      var sheetName = string.Format("Replicated_{0}", this.SelectedSheet.Replace("$", string.Empty));

      ExcelOleDbReader.ExportToXlsx(fileName, results, sheetName);

      Process.Start(fileName);
    }

    #endregion
  }
}