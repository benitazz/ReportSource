using System;

namespace ExcelReportsGenerator.Common.EventArgsHelper
{
  /// <summary>
  /// The close tab event args.
  /// </summary>
  public class CloseTabEventArgs: EventArgs
  {
    #region Public Methods and Operators

    /// <summary>
    /// Gets or sets the close tab.
    /// </summary>
    /// <value>
    /// The close tab.
    /// </value>
    public CloseTab CloseTab { get; set; }

    #endregion
  }
}