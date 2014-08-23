using System;

namespace ExcelReportsGenerator.Common.DependenyInjection
{
  /// <summary>
  /// The import attribute.
  /// </summary>
  [AttributeUsage(AttributeTargets.Property)]
  public class ImportAttribute : Attribute
  {
  }
}