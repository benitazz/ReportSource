#region

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

#endregion

namespace ExcelReportsUtils.Extensions
{
    /// <summary>
    /// The special character helper.
    /// </summary>
    public static class SpecialCharacterRemoverExt
    {
        #region Public Methods and Operators

        /// <summary>
        /// Removes the special characters.
        /// </summary>
        /// <param name="str">
        /// The string.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public static string ReplaceSpecialCharaterWithUnderScore(this string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return null;
            }

            str = str.Trim();

            return Regex.Replace(str, "[^a-zA-Z0-9_]+", "_", RegexOptions.Compiled);
        }

        /// <summary>
        /// Removes the special characters.
        /// </summary>
        /// <param name="str">
        /// The string.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public static string RemoveSpecialCharactersForFileName(this string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return null;
            }

            char[] arr = str.ToCharArray();

            arr = Array.FindAll(arr, c => char.IsLetterOrDigit(c) || char.IsWhiteSpace(c) || c == '-' || c == '_' || c == '(' || c == ')');
            str = new string(arr);

            return str;

            /*str = str.Trim();

            return Regex.Replace(str, "[^a-zA-Z0-9_]+", "_", RegexOptions.Compiled);*/
        }

        /// <summary>
        /// Removes the invalid character for filename.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns>
        /// Returns the valid string for a file name.
        /// </returns>
        public static string RemoveInvalidCharacterForFilename(this string str)
        {
            /*var invalidChars = Path.GetInvalidFileNameChars();

            var invalidCharsRemoved = str.Where(x => !invalidChars.Contains(x)).ToArray();*/

            if (string.IsNullOrEmpty(str))
            {
                return null;
            }

            return Path.GetInvalidFileNameChars().Aggregate(str, (current, c) => current.Replace(c.ToString(CultureInfo.CurrentUICulture), "_"));
        }

        /// <summary>
        /// Removes the special characters.
        /// </summary>
        /// <param name="str">
        /// The string.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public static string RemoveSpecialCharacters(this string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_.]+", "_", RegexOptions.Compiled);
        }

        #endregion
    }
}