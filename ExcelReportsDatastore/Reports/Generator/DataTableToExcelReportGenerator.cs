#region

using System.ComponentModel;

using ExcelReportsDatastore.Reports.Data;

#endregion

namespace ExcelReportsDatastore.Reports.Generator
{
    /// <summary>
    /// The data table to excel report generator.
    /// </summary>
    public class DataTableToExcelReportGenerator : IReportGenerator
    {
        #region Public Methods and Operators

        /// <summary>
        /// The report generator.
        /// </summary>
        /// <param name="reportData">
        /// The report data.
        /// </param>
        /// <param name="reportGeneratorWorker">
        /// The report generator background worker thread.
        /// </param>
        public void GenerateReport(IReportData reportData, BackgroundWorker reportGeneratorWorker)
        {
        }

        #endregion
    }
}