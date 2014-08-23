// --------------------------------------------------------------------------------------------------------------------
// <copyright file="EnumMatchToBooleanConverter.cs" company="BBD">
//   BBD and SARS copyright.
// </copyright>
// <author>Ben Baloyi</author>
// <email>benb@bbd.co.za</email>
// <date>2012-2014</date>
// <summary>
//   The Enum to boolean converter.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

#region

using System;
using System.Globalization;
using System.Windows.Data;

#endregion

namespace ExcelReportsGenerator.Common.Converters
{
  /// <summary>
  ///   The ENUM match to boolean converter.
  /// </summary>
  public class EnumMatchToBooleanConverter : IValueConverter
  {
    #region IValueConverter Members

    /// <summary>
    ///   The ENUM To Boolean Convert
    /// </summary>
    /// <param name="value">
    ///   The value.
    /// </param>
    /// <param name="targetType">
    ///   The target type.
    /// </param>
    /// <param name="parameter">
    ///   The parameter.
    /// </param>
    /// <param name="culture">
    ///   The culture.
    /// </param>
    /// <returns>
    ///   The <see cref="object" />.
    /// </returns>
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
      if ((value == null) || (parameter == null))
      {
        return false;
      }

      var checkValue = value.ToString();
      var target = parameter.ToString();

      return checkValue.Equals(target, StringComparison.InvariantCultureIgnoreCase);
    }

    /// <summary>
    ///   The Boolean to ENUM Convert
    /// </summary>
    /// <param name="value">
    ///   The value.
    /// </param>
    /// <param name="targetType">
    ///   The target type.
    /// </param>
    /// <param name="parameter">
    ///   The parameter.
    /// </param>
    /// <param name="culture">
    ///   The culture.
    /// </param>
    /// <returns>
    ///   The <see cref="object" />.
    /// </returns>
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
      if ((value == null) || (parameter == null))
      {
        return false;
      }

      var useValue = (bool)value;
      var targetValue = parameter.ToString();

      return useValue ? Enum.Parse(targetType, targetValue) : null;
    }

    #endregion
  }
}
