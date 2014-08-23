#region

using System;
using System.Windows.Input;

using BenMVVM;

using ExcelReportsGenerator.Common;
using ExcelReportsGenerator.Common.EventArgsHelper;

#endregion

namespace ExcelReportsGenerator.Models
{
  /// <summary>
  ///   The tab control model.
  /// </summary>
  public class TabControlModel : NotifyBase
  {
    #region Fields

    /// <summary>
    ///   The _header name.
    /// </summary>
    private string _headerName;

    /// <summary>
    ///   The _image source.
    /// </summary>
    private string _imageSource;

    /// <summary>
    ///   The _is active.
    /// </summary>
    private bool _isActive;

    /// <summary>
    ///   The _shell content.
    /// </summary>
    private object _shellContent;

    #endregion

    #region Delegates

    /// <summary>
    ///   The Close Tab delegate.
    /// </summary>
    /// <param name="sender">The sender.</param>
    /// <param name="eventArgs">The <see cref="EventArgs" /> instance containing the event data.</param>
    public delegate void CloseTabDelegete(object sender, CloseTabEventArgs eventArgs);

    #endregion

    #region Public Events

    /// <summary>
    ///   Occurs when [_close tab event].
    /// </summary>
    public event CloseTabDelegete OnCloseTab;

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets the close command.
    /// </summary>
    /// <value>
    ///   The close command.
    /// </value>
    public ICommand CloseCommand
    {
      get
      {
        return new RelayCommand(param => this.CloseApplication());
      }
    }

    /// <summary>
    ///   Gets or sets the header name.
    /// </summary>
    public string HeaderName
    {
      get
      {
        return this._headerName;
      }

      set
      {
        this._headerName = value;
        this.NotifyPropertyChanged();
      }
    }

   /// <summary>
    ///   Gets or sets the image source.
    /// </summary>
    /// <value>
    ///   The image source.
    /// </value>
    public string ImageSource
    {
      get
      {
        return this._imageSource;
      }

      set
      {
        this._imageSource = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets a value indicating whether [is active].
    /// </summary>
    /// <value>
    ///   <c>true</c> if [is active]; otherwise, <c>false</c>.
    /// </value>
    public bool IsActive
    {
      get
      {
        return this._isActive;
      }

      set
      {
        this._isActive = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the shell content.
    /// </summary>
    public object ShellContent
    {
      get
      {
        return this._shellContent;
      }

      set
      {
        this._shellContent = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    /// Gets or sets the _DLL key.
    /// </summary>
    /// <value>
    /// The _DLL key.
    /// </value>
    public string DllKey { get; set; }

    #endregion

    #region Methods

    /// <summary>
    ///   Closes the application.
    /// </summary>
    private void CloseApplication()
    {
      if (this.OnCloseTab != null)
      {
        this.OnCloseTab(this, new CloseTabEventArgs { CloseTab = CloseTab.Close });
      }
    }

    #endregion
  }
}