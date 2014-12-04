#region

using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;

using ExcelReportsDatastore.Helpers;
using ExcelReportsDatastore.Reports.Data;

#endregion

namespace ExcelReportsDatastore.Reports.Generator
{
    /// <summary>
    /// The replicate report generator.
    /// </summary>
    public class ReplicateReportGenerator : IReportGenerator
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

            DataTable replicatedData = null;

            try
            {
                if (reportData == null)
                {
                    return;
                }

                connection.Open();

                var dissectQuery = string.Format("select * from ExcelDataTable");
                
                using (var sqlData = DatabaseReader.GetDatabaseDataTable(connection, dissectQuery))
                {
                    replicatedData = sqlData.Copy();
                    replicatedData.Rows.Clear();

                    var index = 0;

                    foreach (DataRow row in sqlData.Rows)
                    {
                        var value = sqlData.Rows[index++][reportData.ColumnNameFilter];

                        if (value == null)
                        {
                            continue;
                        }

                        var quantity = 0;

                        var strValue = value.ToString();

                        if (string.IsNullOrEmpty(strValue))
                        {
                            continue;
                        }

                        //int.TryParse("123", out quantity);

                        if (!int.TryParse(strValue, out quantity))
                        {
                            AddRowToDatatable(replicatedData, row);

                            continue;
                            /*ExcelReportsUtils.Dialogs.ShowInformation("Only columns that contains numbers can be used to replicate data, please select the correct column and try again");
                            throw new Exception("Could not replicate the data");*/
                        }

                        //var quantity = int.Parse(value.ToString());

                        if (quantity <= 1)
                        {
                            AddRowToDatatable(replicatedData, row);

                            continue;
                        }

                        for (int i = 0; i < quantity; i++)
                        {
                            AddRowToDatatable(replicatedData, row);
                        }                
                    }

                    ExcelWriter.ExportToXlsx(replicatedData, reportData.FileName, reportData.SheetName);
                }
            }
            catch (Exception ex)
            {
                ExcelReportsUtils.Dialogs.ShowError(ex);
            }
            finally
            {
                connection.Dispose();

                if (replicatedData != null)
                {
                    replicatedData.Dispose();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Adds the row to data table.
        /// </summary>
        /// <param name="results">
        /// The results.
        /// </param>
        /// <param name="row">
        /// The row.
        /// </param>
        private static void AddRowToDatatable(DataTable results, DataRow row)
        {
            var newRow = results.NewRow();
            newRow.ItemArray = row.ItemArray;
            results.Rows.Add(newRow);
        }

        #endregion
    }
}