namespace ExcelReportsGenerator.Common
{
  /// <summary>
  ///   The ENUM
  /// </summary>
  public enum ExcelReport
  {
    /// <summary>
    ///   The quantity report.
    /// </summary>
    ReplicateReport, 

    /// <summary>
    ///   The dissect report.
    /// </summary>
    DissectReport
  }

  /// <summary>
  /// The close tab.
  /// </summary>
  public enum CloseTab
  {
    /// <summary>
    /// The close.
    /// </summary>
    Close, 

    /// <summary>
    /// The close all.
    /// </summary>
    CloseAll, 

    /// <summary>
    /// The close all but this.
    /// </summary>
    CloseAllButThis
  }
}