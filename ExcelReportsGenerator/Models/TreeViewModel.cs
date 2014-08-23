#region

using System.Collections.ObjectModel;

using BenMVVM;

#endregion

namespace ExcelReportsGenerator.Models
{
  /// <summary>
  ///   The tree view model.
  /// </summary>
  public class TreeViewModel : NotifyBase
  {
    #region Fields

    /// <summary>
    ///   The _children
    /// </summary>
    private ObservableCollection<TreeViewModel> _children;

    /// <summary>
    /// The _dll key.
    /// </summary>
    private string _dllKey;

    /// <summary>
    ///   The _is expanded.
    /// </summary>
    private bool _isExpanded;

    /// <summary>
    ///   The _is selected.
    /// </summary>
    private bool _isSelected;

    /// <summary>
    ///   The _title.
    /// </summary>
    private string _title;

    /// <summary>
    /// The _image source.
    /// </summary>
    private string _imageSource;

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets or sets the children.
    /// </summary>
    /// <value>
    ///   The children.
    /// </value>
    public ObservableCollection<TreeViewModel> Children
    {
      get
      {
        return this._children;
      }

      set
      {
        this._children = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the DLL key.
    /// </summary>
    /// <value>
    ///   The DLL key.
    /// </value>
    public string DllKey
    {
      get
      {
        return this._dllKey;
      }

      set
      {
        this._dllKey = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets a value indicating whether [is expanded].
    /// </summary>
    /// <value>
    ///   <c>true</c> if [is expanded]; otherwise, <c>false</c>.
    /// </value>
    public bool IsExpanded
    {
      get
      {
        return this._isExpanded;
      }

      set
      {
        this._isExpanded = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets a value indicating whether [is selected].
    /// </summary>
    /// <value>
    ///   <c>true</c> if [is selected]; otherwise, <c>false</c>.
    /// </value>
    public bool IsSelected
    {
      get
      {
        return this._isSelected;
      }

      set
      {
        this._isSelected = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the title.
    /// </summary>
    public string Title
    {
      get
      {
        return this._title;
      }

      set
      {
        this._title = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    /// Gets or sets the image source.
    /// </summary>
    /// <value>
    /// The image source.
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

    #endregion
  }
}