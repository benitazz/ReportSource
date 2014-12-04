#region

using System.ComponentModel;
using System.Data;

using ExcelReportsDatastore.Reports.Data;

#endregion

namespace ExcelReportsDatastore.Reports.Loader
{
    /// <summary>
    /// The ReportLoader interface.
    /// </summary>
    public interface IReportLoader
    {
        #region Public Methods and Operators

        /// <summary>
        /// Loads the report data.
        /// </summary>
        /// <param name="reportData">The report data.</param>
        /// <param name="worker">The worker.</param>
        /// <returns></returns>
        DataTable LoadReportData(IReportData reportData, BackgroundWorker worker);

        #endregion
    }
}