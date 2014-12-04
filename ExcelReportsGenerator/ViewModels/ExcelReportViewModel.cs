#region

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;

using BenMVVM;

using ExcelReportsDatastore.Common;
using ExcelReportsDatastore.Factories;
using ExcelReportsDatastore.Reports.Data;

using ExcelReportsGenerator.Common;
using ExcelReportsGenerator.Common.Helpers;

using ExcelReportsUtils;
using ExcelReportsUtils.Extensions;

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
        /// The _report generator worker
        /// </summary>
        private BackgroundWorker _reportGeneratorWorker;

        /// <summary>
        /// The _background worker.
        /// </summary>
        private BackgroundWorker _reportLoaderWorker;

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
        /// The _is progress bar visible
        /// </summary>
        private bool _isProgressBarVisible;

        /// <summary>
        /// The _progress text
        /// </summary>
        private string _progressText;

        /// <summary>
        /// The _progress value
        /// </summary>
        private int _progressValue;

        /// <summary>
        /// The _report data
        /// </summary>
        private IReportData _reportData;

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
                return new RelayCommand(param => this.ReportGenerator());
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
        /// Gets or sets a value indicating whether this instance is progress bar visible.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is progress bar visible; otherwise, <c>false</c>.
        /// </value>
        public bool IsProgressBarVisible
        {
            get
            {
                return this._isProgressBarVisible;
            }

            set
            {
                this._isProgressBarVisible = value;
                this.NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Gets or sets the progress text.
        /// </summary>
        /// <value>
        /// The progress text.
        /// </value>
        public string ProgressText
        {
            get
            {
                return this._progressText;
            }

            set
            {
                this._progressText = value;
                this.NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Gets or sets the progress value.
        /// </summary>
        /// <value>
        /// The progress value.
        /// </value>
        public int ProgressValue
        {
            get
            {
                return this._progressValue;
            }

            set
            {
                this._progressValue = value;
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
        ///   Files the browser handler.
        /// </summary>
        public void FileBrowserHandler()
        {
            try
            {
                this.IsBusy = true;
                this.ProgressText = string.Empty;
                this.IsProgressBarVisible = false;

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

                var filename = Path.GetFileName(this.FileName);

                this.ProgressText = string.Format("Loading Data From {0}...", filename);

                this._reportLoaderWorker = new BackgroundWorker();
                this._reportLoaderWorker.DoWork += this.ReportLoaderWorkerOnDoWork;
                this._reportLoaderWorker.RunWorkerCompleted += this.ReportLoaderWorkerRunWorkerCompleted;
                this._reportLoaderWorker.ProgressChanged += this.ReportLoaderWorkerProgressChanged;
                this._reportLoaderWorker.WorkerReportsProgress = true;
                this._reportLoaderWorker.WorkerSupportsCancellation = true;

                if (fileExtension.Contains(".xl"))
                {
                    // this._reportLoaderWorker.RunWorkerAsync(this.FileName);
                    // ExcelOleDbReader.GetExcelDataTable(this.FileName, this._reportLoaderWorker);

                    this._reportData = new ReportData { FileName = this.FileName, };

                    this._reportLoaderWorker.RunWorkerAsync(this._reportData);

                    /*this._excelSheetDataTable = ExcelOleDbReader.GetExcelDataTable(this.FileName);

                    this.ExcelData = this._excelSheetDataTable.DefaultView;*/

                    /*this.TotalRecords = this.ExcelData.Count;

                    this.ColumnsNamesList = new ObservableCollection<string>();

                    foreach (DataColumn column in this._excelSheetDataTable.Columns)
                    {
                        this.ColumnsNamesList.Add(column.ColumnNameFilter);
                    }

                    this.NotifyPropertyChanged(() => this.ColumnsNamesList);
                    this.IsColumnFilterEnabled = false;*/

                    /*this.Sheets = ExcelOleDbReader.GetSheetNames();

                    this.SelectedSheet = this.Sheets[0];*/
                    // this.SelectedColumnFilter = "Quantity";
                }
            }
            catch (Exception exception)
            {
                Dialogs.ShowError(exception);
            }

            this.IsBusy = false;
        }

        /// <summary>
        ///   Quantities the report.
        /// </summary>
        public void ReportGenerator()
        {
            if (this._excelSheetDataTable == null)
            {
                return;
            }

            this.IsBusy = true;
            DataTable results = this._excelSheetDataTable.Copy();
            results.Rows.Clear();

            this._reportGeneratorWorker = new BackgroundWorker();
            this._reportGeneratorWorker.DoWork += this.ReportGeneratorWorkerOnDoWork;
            this._reportGeneratorWorker.RunWorkerCompleted += this.ReportGeneratorWorkerRunWorkerCompleted;
            this._reportGeneratorWorker.ProgressChanged += this.ReportGeneratorWorkerOnProgressChanged;
            this._reportGeneratorWorker.WorkerReportsProgress = true;
            this._reportGeneratorWorker.WorkerSupportsCancellation = true;

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
        /// Handles the ProgressChanged event of the _reportLoaderWorker control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event.
        /// </param>
        /// <param name="e">
        /// The <see cref="ProgressChangedEventArgs"/> instance containing the event data.
        /// </param>
        /// <exception cref="System.NotImplementedException">
        /// </exception>
        public void ReportLoaderWorkerProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.IsBusy = true;

            /*if (!this.IsProgressBarVisible)
            {
                this.IsProgressBarVisible = true;
            }

            this.ProgressValue = e.ProgressPercentage;

            this.NotifyPropertyChanged(() => this.ProgressValue);*/

            // Thread.Sleep(5);
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the _reportLoaderWorker control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event.
        /// </param>
        /// <param name="e">
        /// The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.
        /// </param>
        /// <exception cref="System.NotImplementedException">
        /// </exception>
        public void ReportLoaderWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                this.IsBusy = false;
                return;
            }

            if (e.Error != null)
            {
                MessageBox.Show("Error while performing the data load", "Data Load");
                this.IsBusy = false;
                return;
            }

            var dissectHelper = e.Result as IReportData;

            if (dissectHelper != null)
            {
                Process.Start(dissectHelper.Directory);
                this.IsBusy = false;
                return;
            }

            var dataTable = e.Result as DataTable;

            if (dataTable == null)
            {
                return;
            }

            this._excelSheetDataTable = dataTable;
            this.ExcelData = this._excelSheetDataTable.DefaultView;
            this.TotalRecords = this.ExcelData.Count;

            this.ColumnsNamesList = new ObservableCollection<string>();

            foreach (DataColumn column in this._excelSheetDataTable.Columns)
            {
                this.ColumnsNamesList.Add(column.ColumnName);
            }

            this.NotifyPropertyChanged(() => this.ColumnsNamesList);
            this.IsColumnFilterEnabled = false;

            this.IsBusy = false;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Excels the dissect report generator.
        /// </summary>
        /// <param name="results">
        /// The results.
        /// </param>
        private void ExcelDissectReportGenerator(DataTable results)
        {
            if (string.IsNullOrEmpty(this.SelectedColumnFilter))
            {
                Dialogs.ShowWarning("Please select a column to be used for dissecting.");
                return;
            }

            var fileName = Path.GetFileNameWithoutExtension(this.FileName);

            fileName = fileName.RemoveSpecialCharacters();

            var datetime = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss-tt");

            var directoryName = string.Format(@"C:\Dissect\{0} {1}", fileName, datetime);

            FileBrowserHelper.CreateDirectory(directoryName);

            this._reportData = new ReportData
                                   {
                                       ColumnNameFilter = this.SelectedColumnFilter, 
                                       FileName = this.FileName, 
                                       Directory = directoryName, 
                                       ReportGeneratorType = ReportGeneratorType.DissectData
                                   };

            this.ProgressText = string.Format("Dissecting Data using {0}...", this.SelectedColumnFilter);

            this._reportGeneratorWorker.RunWorkerAsync(this._reportData);
        }

        /// <summary>
        /// Excels the replicate report generator.
        /// </summary>
        /// <param name="results">
        /// The results.
        /// </param>
        private void ExcelReplicateReportGenerator(DataTable results)
        {
            /*var index = 0;

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
            }*/
            this.ProgressText = string.Empty;

            if (string.IsNullOrEmpty(this.SelectedColumnFilter))
            {
                Dialogs.ShowWarning("Please select the column name to be used for filtering data.");
                return;
            }

            this.ProgressText = string.Format("Replicating Data using {0}...", this.SelectedColumnFilter);

            this._reportData = new ReportData
                                   {
                                       ColumnNameFilter = this.SelectedColumnFilter, 
                                       FileName = @"c:\test.xlsx", 
                                       ReportGeneratorType = ReportGeneratorType.ReplicateData, 
                                       SheetName = string.Format("Replicated_{0}", this.SelectedColumnFilter)
                                   };

            this._reportGeneratorWorker.RunWorkerAsync(this._reportData);

            /*// var fileName = string.Format("{0}_{1}", DateTime.Now, Path.GetFileName(this.FileName));
            var fileName = @"c:\test.xlsx";

            // var sheetName = string.Format("Replicated_{0}", this.SelectedSheet.Replace("$", string.Empty));

            var sheetName = "Sheet1";

            ExcelWriter.ExportToXlsx(results, fileName, sheetName);

            Process.Start(fileName);*/
        }

        /// <summary>
        /// The report generator worker on do work.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="doWorkEventArgs">
        /// The do work event args.
        /// </param>
        private void ReportGeneratorWorkerOnDoWork(object sender, DoWorkEventArgs doWorkEventArgs)
        {
            var reportData = doWorkEventArgs.Argument as IReportData;

            if (reportData == null)
            {
                return;
            }

            var reportGenerator = GenerateReportFactory.GetReportGeneratorInstance(reportData.ReportGeneratorType);

            reportGenerator.GenerateReport(reportData, this._reportGeneratorWorker);
        }

        /// <summary>
        /// Reports the generator worker on progress changed.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="progressChangedEventArgs">
        /// The <see cref="ProgressChangedEventArgs"/> instance containing the event data.
        /// </param>
        private void ReportGeneratorWorkerOnProgressChanged(
            object sender, 
            ProgressChangedEventArgs progressChangedEventArgs)
        {
            this.IsBusy = true;
        }

        /// <summary>
        /// Reports the generator worker run worker completed.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.
        /// </param>
        private void ReportGeneratorWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                this.IsBusy = false;
                return;
            }

            if (e.Error != null)
            {
                // Dialogs.ShowError("Error while performing the data load", "Data Load");
                this.IsBusy = false;
                return;
            }

            switch (this._reportData.ReportGeneratorType)
            {
                case ReportGeneratorType.DissectData:
                    {
                        Process.Start(this._reportData.Directory);
                        this.IsBusy = false;
                        return;
                    }

                case ReportGeneratorType.ReplicateData:
                    {
                        Process.Start(this._reportData.FileName);
                        this.IsBusy = false;
                        return;
                    }
            }
        }

        /// <summary>
        /// Reports the loader worker on do work.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="doWorkEventArgs">
        /// The <see cref="DoWorkEventArgs"/> instance containing the event data.
        /// </param>
        private void ReportLoaderWorkerOnDoWork(object sender, DoWorkEventArgs doWorkEventArgs)
        {
            var reportData = doWorkEventArgs.Argument as IReportData;

            if (reportData == null)
            {
                return;
            }

            var reportLoader = LoadReportFactory.GetReportLoaderInstance(reportData.ReportLoaderType);

            doWorkEventArgs.Result = reportLoader.LoadReportData(reportData, this._reportLoaderWorker);
         }

        #endregion
    }
}