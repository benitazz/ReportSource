#region

using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

#endregion

namespace ExcelReportsGenerator.Common.WpfExtensions
{
  /// <summary>
  /// The tree view extension.
  /// </summary>
  internal class TreeViewExtension : TreeView, INotifyPropertyChanged
  {
    #region Static Fields

    /// <summary>
    ///   The selected items property
    /// </summary>
    public static readonly DependencyProperty SelectedItemsProperty = DependencyProperty.Register(
      "SelectedItem", 
      typeof(object), 
      typeof(TreeViewExtension), 
      new PropertyMetadata(null));

    #endregion

    #region Constructors and Destructors

    /// <summary>
    ///   Initializes a new instance of the <see cref="TreeViewExtension" /> class.
    /// </summary>
    public TreeViewExtension()
    {
      this.SelectedItemChanged += this.TreeViewSelectedItemChanged;
    }

    #endregion

    #region Public Events

    /// <summary>
    ///   Occurs when [property changed].
    /// </summary>
    public event PropertyChangedEventHandler PropertyChanged;

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets the selected item in a <see cref="T:System.Windows.Controls.TreeView" />. This is a dependency property.
    /// </summary>
    /// <returns>
    ///   The selected object in the <see cref="T:System.Windows.Controls.TreeView" />, or null if no item is selected. The
    ///   default value is null.
    /// </returns>
    public new object SelectedItem
    {
      get
      {
        return this.GetValue(SelectedItemProperty);
      }

      set
      {
        this.SetValue(SelectedItemsProperty, value);
        this.NotifyPropertyChanged("SelectedItem");
      }
    }

    #endregion

    #region Methods

    /// <summary>
    /// Notifies the property changed.
    /// </summary>
    /// <param name="aPropertyName">
    /// Name of the aggregate property.
    /// </param>
    private void NotifyPropertyChanged(string aPropertyName)
    {
      if (this.PropertyChanged != null)
      {
        this.PropertyChanged(this, new PropertyChangedEventArgs(aPropertyName));
      }
    }

    /// <summary>
    /// The TreeView selected item changed.
    /// </summary>
    /// <param name="sender">
    /// The sender.
    /// </param>
    /// <param name="e">
    /// The decimal.
    /// </param>
    private void TreeViewSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
      this.SelectedItem = base.SelectedItem;
    }

    #endregion
  }
}