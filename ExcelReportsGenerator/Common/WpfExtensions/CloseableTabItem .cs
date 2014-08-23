#region

using System.Windows;
using System.Windows.Controls;

#endregion

namespace ExcelReportsGenerator.Common.WpfExtensions
{
  /// <summary>
  /// The closeable tab item.
  /// </summary>
  public class CloseableTabItem : TabItem
  {
    #region Constructors and Destructors

    /// <summary>
    /// Initializes static members of the <see cref="CloseableTabItem"/> class.
    /// </summary>
    static CloseableTabItem()
    {
      DefaultStyleKeyProperty.OverrideMetadata(
        typeof(CloseableTabItem), 
        new FrameworkPropertyMetadata(typeof(CloseableTabItem)));
    }

    #endregion
  }
}