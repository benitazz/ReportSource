#region

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using ExcelReportsDatastore;

using ExcelReportsGenerator.Common;
using ExcelReportsGenerator.Common.Helpers;

using Microsoft.Win32;

#endregion

namespace ExcelReportsGenerator.ViewModels
{
    /// <summary>
    ///     The text report view model.
    /// </summary>
    public class TextReportViewModel : ViewModelBase, IShellViewModel
    {
        #region Constants

        /// <summary>
        ///     The default text filters.
        /// </summary>
        private const string DefaultFilters = "GJ;AP;CL";

        #endregion

        #region Fields

        /// <summary>
        ///     The _data array length.
        /// </summary>
        private int _dataArrayLength;

        /// <summary>
        ///     The _data filter.
        /// </summary>
        private string _dataFilter;

        /// <summary>
        ///     The _data table.
        /// </summary>
        private DataTable _dataTable;

        /// <summary>
        ///     The _is busy.
        /// </summary>
        private bool _isBusy;

        private DataView _textData;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        ///     Initializes a new instance of the <see cref="TextReportViewModel" /> class.
        /// </summary>
        public TextReportViewModel()
        {
            this.IsBusy = false;

            this.DataFilter = DefaultFilters;
        }

        #endregion

        #region Public Properties

        /// <summary>
        ///     Gets or sets the data filter.
        /// </summary>
        /// <value>
        ///     The data filter.
        /// </value>
        public string DataFilter
        {
            get
            {
                return this._dataFilter;
            }

            set
            {
                this._dataFilter = value;
                this.NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Gets or sets the text data.
        /// </summary>
        /// <value>
        /// The text data.
        /// </value>
        public DataView TextData
        {
            get
            {
                return this._textData;
            }

            set
            {
                this._textData = value;
                this.NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether is busy.
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

        #endregion

        #region Public Methods and Operators

        /// <summary>
        ///     Files the browser handler.
        /// </summary>
        public void FileBrowserHandler()
        {
            this.IsBusy = true;

            var fileDialog = new OpenFileDialog { Filter = "Text Files (*.txt)|*.txt|All Files(*.*)|*.*" };

            var fileName = FileBrowserHelper.GetfileName(fileDialog);

            if (string.IsNullOrEmpty(fileName))
            {
                this.IsBusy = false;
                return;
            }

            using (var reader = new StreamReader(File.OpenRead(fileName)))
            {
                var filters = this.DataFilter.Split(';');

                this.ProcessTextFile(reader, filters);
            }

            if (this._dataFilter != null)
            {
                var defaultView  = this._dataTable.DefaultView;
                defaultView.Sort = "Column1 ASC, Column3 ASC";
                this.TextData = defaultView;
            }

            this.IsBusy = false;
        }

        /// <summary>
        ///     Reports the generator.
        /// </summary>
        public void ReportGenerator()
        {
           if (this._dataTable == null)
            {
                return;
            }

            this.IsBusy = true;

            // var fileName = string.Format("{0}_{1}", DateTime.Now, Path.GetFileName(this.FileName));
            var fileName = @"c:\test.xlsx";

            var sheetName = string.Format("Sheet1");

            ExcelWriter.ExportToXlsx(fileName, this._dataTable, sheetName);

            Process.Start(fileName);

            this.IsBusy = false;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Adds the row.
        /// </summary>
        /// <param name="lineValues">
        /// The line values.
        /// </param>
        /// <param name="table">
        /// The table.
        /// </param>
        private static void AddRow(string[] lineValues, DataTable table)
        {
            var row = table.NewRow();

            for (var i = 0; i < lineValues.Length; i++)
            {
                row[i] = lineValues[i];
            }

            table.Rows.Add(row);
        }

        /// <summary>
        /// Adds the text values to array.
        /// </summary>
        /// <param name="index">
        /// The index.
        /// </param>
        /// <param name="values">
        /// The values.
        /// </param>
        /// <param name="lineValue">
        /// The line value.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        private static int AddTextValuesToArray(int index, List<string> values, string lineValue)
        {
            if (index == 1 || index == 2)
            {
                values.Add(lineValue);
                ++index;
                return index;
            }

            if (Regex.IsMatch(lineValue, @"^\d+$"))
            {
                values.Add(lineValue);
                ++index;
                return index;
            }

            if (Regex.IsMatch(lineValue, @"^\d{2}/\d{2}/\d{4}"))
            {
                values.Add(lineValue);
                ++index;

                // DateTime date = DateTime.ParseExact(lineValue.Trim(), "dd/MM/yyyy", null);
                return index;
            }

            if (Regex.IsMatch(lineValue, @"^\d*?\.\d*?$") || Regex.IsMatch(lineValue, @"^\d*?,\d*\.\d*?$"))
            {
                values.Add(lineValue);
                ++index;
                return index;
            }

            var prevIndex = values.Count - 1;

            var prevValue = values[prevIndex];

            if (!Regex.IsMatch(prevValue, @"^\d+$") && !Regex.IsMatch(lineValue, @"^\d{2}/\d{2}/\d{4}")
                && !Regex.IsMatch(lineValue, @"^\d*?\.\d*?$") && !Regex.IsMatch(lineValue, @"^\d*?,\d*\.\d*?$"))
            {
                values[prevIndex] = string.Format("{0} {1}", prevValue, lineValue);
                ++index;
                return index;
            }

            values.Add(lineValue);
            ++index;

            return index;
        }

        /// <summary>
        /// Gets the table.
        /// </summary>
        /// <param name="lineValues">
        /// The line values.
        /// </param>
        /// <returns>
        /// Returns the data table.
        /// </returns>
        private DataTable GetTable(string[] lineValues)
        {
            // Here we create a DataTable with four columns.
            var table = new DataTable();

            var index = 1;

            for (int i = 1; i <= lineValues.Length; i++)
            {
                table.Columns.Add(string.Format("Column{0}", index), typeof(string));
                ++index;
            }

            AddRow(lineValues, table);

            return table;
        }

        /// <summary>
        /// Processes the text file.
        /// </summary>
        /// <param name="reader">
        /// The reader.
        /// </param>
        /// <param name="filters">
        /// The filters.
        /// </param>
        private void ProcessTextFile(StreamReader reader, string[] filters)
        {
            var lineIndex = 1;

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();

                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }

                try
                {
                    line = line.Trim(); // Removes trailing spaces.

                    var isFilterExistsInLine = filters.Any(filter => line.ToUpper().StartsWith(filter.Trim()));

                    if (!isFilterExistsInLine)
                    {
                        continue;
                    }

                    // var filesvalues = System.Text.RegularExpressions.Regex.Split(line, @"\s{2,}");
                    var tempValues = line.Split(' ');

                    var values = new List<string>();

                    var index = 1;

                    foreach (var lineValue in tempValues.Where(val => !string.IsNullOrEmpty(val)))
                    {
                        index = AddTextValuesToArray(index, values, lineValue);
                    }

                    if (lineIndex == 1)
                    {
                        this._dataTable = this.GetTable(values.ToArray());
                        this._dataArrayLength = values.Count;

                        ++lineIndex;
                        continue;
                    }

                    if (values.Count == this._dataArrayLength)
                    {
                        AddRow(values.ToArray(), this._dataTable);
                    }

                    ++lineIndex;
                }
                catch (Exception exception)
                {
                    throw new Exception(exception.Message);
                }
            }
        }

        #endregion
    }
}