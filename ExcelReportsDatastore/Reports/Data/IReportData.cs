// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IReportData.cs" company="">
//   
// </copyright>
// <summary>
//   The DissectHelper interface.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using ExcelReportsDatastore.Common;

namespace ExcelReportsDatastore.Reports.Data
{
    /// <summary>
    /// The DissectHelper interface.
    /// </summary>
    public interface IReportData
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the column name.
        /// </summary>
        string ColumnNameFilter { get; set; }

        /// <summary>
        /// Gets or sets the directory.
        /// </summary>
        /// <value>
        /// The directory.
        /// </value>
        string Directory { get; set; }

        /// <summary>
        /// Gets or sets the name of the file.
        /// </summary>
        /// <value>
        /// The name of the file.
        /// </value>
        string FileName { get; set; }

        /// <summary>
        /// Gets or sets the table name.
        /// </summary>
        string SheetName { get; set; }

        /// <summary>
        /// Gets or sets the type of the report generator.
        /// </summary>
        /// <value>
        /// The type of the report generator.
        /// </value>
        ReportGeneratorType ReportGeneratorType { get; set; }

        /// <summary>
        /// Gets or sets the type of the report loader.
        /// </summary>
        /// <value>
        /// The type of the report loader.
        /// </value>
        ReportLoaderType ReportLoaderType { get; set; }

        #endregion
    }
}