#region

using ExcelReportsDatastore.Common;
using ExcelReportsDatastore.Reports.Loader;

#endregion

namespace ExcelReportsDatastore.Factories
{
    /// <summary>
    /// The load report factory.
    /// </summary>
    public class LoadReportFactory
    {
        #region Public Methods and Operators

        /// <summary>
        /// The get report loader instance.
        /// </summary>
        /// <param name="reportLoaderType">
        /// The report loader type.
        /// </param>
        /// <returns>
        /// The <see cref="IReportLoader"/>.
        /// </returns>
        public static IReportLoader GetReportLoaderInstance(ReportLoaderType reportLoaderType)
        {
            switch (reportLoaderType)
            {
                case ReportLoaderType.ExcelReport:
                    {
                        return new ExcelReportDataLoader();
                    }

                case ReportLoaderType.TextReport:
                    {
                        return new TextReportDataLoader();
                    }
            }

            return null;
        }

        #endregion
    }
}