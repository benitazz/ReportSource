namespace ExcelReportsGenerator.ViewModels
{
  /// <summary>
  /// The text report view model.
  /// </summary>
  public class TextReportViewModel : ViewModelBase, IShellViewModel
  {
    #region Member Variables

    /// <summary>
    /// The _is busy.
    /// </summary>
    private bool _isBusy;

    #endregion

    #region Constructor

    /// <summary>
    /// Initializes a new instance of the <see cref="TextReportViewModel"/> class.
    /// </summary>
    public TextReportViewModel()
    {
      this.IsBusy = false;
    }

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets or sets a value indicating whether is busy.
    /// </summary>
    public bool IsBusy
    {
      get
      {
        return this._isBusy;
      }
      set
      {
        this._isBusy = value;
        this.NotifyPropertyChanged();
      }
    }

    #endregion

    #region Public Methods and Operators

    /// <summary>
    ///   Reports the generator.
    /// </summary>
    public void ExcelReportGenerator()
    {
    }

    /// <summary>
    ///   Files the browser handler.
    /// </summary>
    public void FileBrowserHandler()
    {
    }

    #endregion
  }
}