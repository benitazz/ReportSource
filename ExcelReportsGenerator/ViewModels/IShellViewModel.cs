namespace ExcelReportsGenerator.ViewModels
{
  /// <summary>
  /// The ShellViewModel interface.
  /// </summary>
  public interface IShellViewModel
  {
    #region Public Properties

    /// <summary>
    /// Gets or sets a value indicating whether is busy.
    /// </summary>
    bool IsBusy { get; set; }

    /// <summary>
    /// Files the browser handler.
    /// </summary>
    void FileBrowserHandler();

    /// <summary>
    /// Reports the generator.
    /// </summary>
    void ReportGenerator();

    #endregion
  }
}