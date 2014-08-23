#region

using System;
using System.Globalization;
using System.Windows.Data;

#endregion

namespace ExcelReportsGenerator.Common.Converters
{
  /// <summary>
  ///   The boolean to object converter.
  /// </summary>
  public sealed class BooleanToObjectConverter : IValueConverter
  {
    #region Constructors and Destructors

    /// <summary>
    ///   Initializes a new instance of the <see cref="BooleanToObjectConverter" /> class.
    /// </summary>
    public BooleanToObjectConverter()
    {
      this.Invert = false;
      this.TrueStatus = Binding.DoNothing;
      this.FalseStatus = Binding.DoNothing;
    }

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets or sets the false status.
    /// </summary>
    public object FalseStatus { get; set; }

    /// <summary>
    ///   Gets or sets a value indicating whether invert.
    /// </summary>
    public bool Invert { get; set; }

    /// <summary>
    ///   Gets or sets the true status.
    /// </summary>
    public object TrueStatus { get; set; }

    #endregion

    #region Public Methods and Operators

    /// <summary>
    /// The boolean to object Convert
    /// </summary>
    /// <param name="value">
    /// The value.
    /// </param>
    /// <param name="targetType">
    /// The target type.
    /// </param>
    /// <param name="parameter">
    /// The parameter.
    /// </param>
    /// <param name="culture">
    /// The culture.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
      if (value != null)
      {
        return (bool)value
                 ? (!this.Invert ? this.TrueStatus : this.FalseStatus)
                 : (!this.Invert ? this.FalseStatus : this.TrueStatus);
      }

      return !this.Invert ? this.FalseStatus : this.TrueStatus;
    }

    /// <summary>
    /// The convert back.
    /// </summary>
    /// <param name="value">
    /// The value.
    /// </param>
    /// <param name="targetType">
    /// The target type.
    /// </param>
    /// <param name="parameter">
    /// The parameter.
    /// </param>
    /// <param name="culture">
    /// The culture.
    /// </param>
    /// <returns>
    /// The <see cref="object"/>.
    /// </returns>
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
      if (value != null)
      {
        return !this.Invert ? value.Equals(this.TrueStatus) : !value.Equals(this.TrueStatus);
      }

      return !this.Invert;
    }

    #endregion
  }
}