#region

using ExcelReportsDatastore.Common;
using ExcelReportsDatastore.Reports.Generator;

#endregion

namespace ExcelReportsDatastore.Factories
{
    /// <summary>
    /// The generate report factory.
    /// </summary>
    public class GenerateReportFactory
    {
        #region Public Methods and Operators

        /// <summary>
        /// The get generator instance.
        /// </summary>
        /// <param name="reportGeneratorType">
        /// The report generator type.
        /// </param>
        /// <returns>
        /// The <see cref="IReportGenerator"/>.
        /// </returns>
        public static IReportGenerator GetReportGeneratorInstance(ReportGeneratorType reportGeneratorType)
        {
            switch (reportGeneratorType)
            {
                case ReportGeneratorType.DissectData:
                    {
                        return new DissectReportGenerator();
                    }

                case ReportGeneratorType.ReplicateData:
                    {
                        return new ReplicateReportGenerator();
                    }

                case ReportGeneratorType.DataTableToExcelReport:
                    {
                        return new DataTableToExcelReportGenerator();
                    }
            }

            return null;
        }

        #endregion
    }
}