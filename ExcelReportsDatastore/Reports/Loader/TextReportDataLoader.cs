using System.ComponentModel;
using System.Data;

using ExcelReportsDatastore.Reports.Data;

namespace ExcelReportsDatastore.Reports.Loader
{
    /// <summary>
    /// The text report data loader.
    /// </summary>
    public class TextReportDataLoader : IReportLoader
    {
        #region Implementation of IReportLoader

        /// <summary>
        /// Loads the report data.
        /// </summary>
        /// <param name="reportData">The report data.</param>
        /// <param name="worker">The worker.</param>
        public DataTable LoadReportData(IReportData reportData, BackgroundWorker worker)
        {
            return null;
        }

        #endregion
    }
}