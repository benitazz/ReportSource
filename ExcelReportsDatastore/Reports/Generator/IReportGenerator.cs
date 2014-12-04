#region

using System.ComponentModel;

using ExcelReportsDatastore.Reports.Data;

#endregion

namespace ExcelReportsDatastore.Reports.Generator
{
    /// <summary>
    /// The ReportGenerator interface.
    /// </summary>
    public interface IReportGenerator
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
        void GenerateReport(IReportData reportData, BackgroundWorker reportGeneratorWorker);

        #endregion
    }
}