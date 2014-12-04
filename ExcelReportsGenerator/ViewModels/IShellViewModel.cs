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
    /// Gets or sets the progress text.
    /// </summary>
    /// <value>
    /// The progress text.
    /// </value>
    string ProgressText { get; set; }

    /// <summary>
    /// Gets or sets the progress value.
    /// </summary>
    /// <value>
    /// The progress value.
    /// </value>
    int ProgressValue { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether this instance is progress bar visible.
    /// </summary>
    /// <value>
    /// <c>true</c> if this instance is progress bar visible; otherwise, <c>false</c>.
    /// </value>
    bool IsProgressBarVisible { get; set; }

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