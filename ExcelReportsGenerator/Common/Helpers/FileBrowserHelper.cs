#region

using System.IO;

using Microsoft.Win32;

#endregion

namespace ExcelReportsGenerator.Common.Helpers
{
  /// <summary>
  /// The file browser helper.
  /// </summary>
  public class FileBrowserHelper
  {
    #region Public Methods and Operators

    /// <summary>
    /// The create directory.
    /// </summary>
    /// <param name="directory">
    /// The directory.
    /// </param>
    public static void CreateDirectory(string directory)
    {
      if (!Directory.Exists(directory))
      {
        Directory.CreateDirectory(directory);
      }
    }

    /// <summary>
    /// The create file.
    /// </summary>
    /// <param name="directory">
    /// The directory.
    /// </param>
    /// <param name="outputfileName">
    /// The output file name.
    /// </param>
    /// <returns>
    /// The <see cref="string"/>.
    /// </returns>
    public static string CreateFile(string directory, string outputfileName)
    {
      var filePath = Path.Combine(directory, outputfileName);

      if (!File.Exists(filePath))
      {
        File.Create(filePath).Close();
      }
      else
      {
        File.WriteAllText(filePath, string.Empty); // Clear the content of the file
      }

      return filePath;
    }

    /// <summary>
    /// The get file name from the directory.
    /// </summary>
    /// <returns>
    /// The <see cref="string"/>.
    /// </returns>
    public static string GetfileName()
    {
      var fileBrowser = new OpenFileDialog { Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx" };

      var isFileLocated = fileBrowser.ShowDialog();

      return isFileLocated.Value != true ? string.Empty : fileBrowser.FileName;
    }

    /// <summary>
    /// Get file name.
    /// </summary>
    /// <param name="fileDialog">The file dialog.</param>
    /// <returns>
    /// Returns the file Path.
    /// </returns>
    public static string GetfileName(OpenFileDialog fileDialog)
    {
      if (fileDialog == null)
      {
        fileDialog = new OpenFileDialog();
      }

      var isFileLocated = fileDialog.ShowDialog();

      return isFileLocated.Value != true ? string.Empty : fileDialog.FileName;
    }

    #endregion
  }
}