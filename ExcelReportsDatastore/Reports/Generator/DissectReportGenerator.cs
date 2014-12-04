#region

using System;
using System.ComponentModel;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.IO;

using ExcelReportsDatastore.Reports.Data;

using ExcelReportsUtils.Extensions;

#endregion

namespace ExcelReportsDatastore.Reports.Generator
{
    /// <summary>
    /// The dissect report generator.
    /// </summary>
    public class DissectReportGenerator : IReportGenerator
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
            reportGeneratorWorker.ReportProgress(1);
            var connection = ExcelOleDbReader.GetLocalConnection();

            try
            {
                if (reportData == null)
                {
                    return;
                }

                connection.Open();

                var selectDisctinctDissectQuery = string.Format(
                    "select DISTINCT {0} from {1}", 
                    reportData.ColumnNameFilter, 
                    "ExcelDataTable");

                using (var cmd = new SqlCeCommand(selectDisctinctDissectQuery, connection))
                {
                    using (var dissectColumnReader = cmd.ExecuteReader())
                    {
                        while (dissectColumnReader.Read())
                        {
                            var dissectValue = dissectColumnReader.GetString(0);

                            if (string.IsNullOrEmpty(dissectValue))
                            {
                                continue;
                            }

                            var newValue = dissectValue.Replace("'", "''");

                            var dissectQuery = string.Format(
                                "select * from {0} Where [{1}]='{2}'", 
                                "ExcelDataTable", 
                                reportData.ColumnNameFilter, 
                                newValue);

                            var dataTable = DatabaseReader.GetDatabaseDataTable(connection, dissectQuery);

                            dissectValue = dissectValue.RemoveInvalidCharacterForFilename();

                            var filename = string.Format(
                                @"{0}\{1}_{2}", 
                                reportData.Directory, 
                                dissectValue, 
                                Path.GetFileName(reportData.FileName));

                            var sheetName = dissectValue;

                            // var filename = string.Format(@"{0}\{1}.{2}", directory, dissectValue, Path.GetExtension(fileName));
                            if (sheetName.Length > 30)
                            {
                                sheetName = sheetName.Substring(0, 30);
                            }

                            //ExcelWriter.ExportBigDataToXlsx(dataTable, filename, sheetName);

                            if (dataTable.Rows.Count > 10000)
                            {
                                ExcelWriter.ExportBigDataToXlsx(dataTable, filename, sheetName);
                            }
                            else
                            {
                                ExcelWriter.ExportToXlsx(dataTable, filename, sheetName);
                            }
                        }
                    }
                }
             }
            catch (Exception ex)
            {
               ExcelReportsUtils.Dialogs.ShowError(ex);
            }
            finally
            {
                connection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion
    }
}