namespace ExcelReportsDatastore.Common
{
    /// <summary>
    /// The report operation.
    /// </summary>
    public enum ReportOperation
    {
        /// <summary>
        /// The load data.
        /// </summary>
        LoadData, 

        /// <summary>
        /// The dissect data.
        /// </summary>
        DissectData, 

        /// <summary>
        /// The replicate data.
        /// </summary>
        ReplicateData
    }

    /// <summary>
    /// The report generator type.
    /// </summary>
    public enum ReportGeneratorType
    {
        /// <summary>
        /// The dissect data.
        /// </summary>
        DissectData, 

        /// <summary>
        /// The replicate data.
        /// </summary>
        ReplicateData,

        /// <summary>
        /// The data table to excel report
        /// </summary>
        DataTableToExcelReport
    }

    /// <summary>
    /// The report loader type.
    /// </summary>
    public enum ReportLoaderType
    {
        /// <summary>
        /// The dissect data.
        /// </summary>
        ExcelReport, 

        /// <summary>
        /// The replicate data.
        /// </summary>
        TextReport
    }
}