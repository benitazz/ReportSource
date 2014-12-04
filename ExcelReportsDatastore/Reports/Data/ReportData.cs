// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ReportData.cs" company="">
//   
// </copyright>
// <summary>
//   The dissect helper.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using ExcelReportsDatastore.Common;

namespace ExcelReportsDatastore.Reports.Data
{
    /// <summary>
    /// The dissect helper.
    /// </summary>
    public class ReportData : IReportData
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the column name.
        /// </summary>
        public string ColumnNameFilter { get; set; }

        /// <summary>
        /// Gets or sets the directory.
        /// </summary>
        /// <value>
        /// The directory.
        /// </value>
        public string Directory { get; set; }

        /// <summary>
        /// Gets or sets the name of the file.
        /// </summary>
        /// <value>
        /// The name of the file.
        /// </value>
        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets the table name.
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets or sets the type of the report generator.
        /// </summary>
        /// <value>
        /// The type of the report generator.
        /// </value>
        public ReportGeneratorType ReportGeneratorType { get; set; }

        /// <summary>
        /// Gets or sets the type of the report loader.
        /// </summary>
        /// <value>
        /// The type of the report loader.
        /// </value>
        public ReportLoaderType ReportLoaderType { get; set; }

        /// <summary>
        /// Gets or sets the report operation.
        /// </summary>
        /// <value>
        /// The report operation.
        /// </value>
        public ReportOperation ReportOperation { get; set; }

        #endregion
    }
}