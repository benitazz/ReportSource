#region

using System;
using System.Configuration;

#endregion

namespace ExcelReportGeneratorDataProvider
{
    /// <summary>
    /// The connection string helper.
    /// </summary>
    public class ConnectionStringHelper
    {
        #region Public Methods and Operators

        /// <summary>
        /// The get connection string.
        /// </summary>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public static string GetConnectionString()
        {
            var connectionString = ConfigurationManager.ConnectionStrings["ExcelReportsDbEntities"];

            if (connectionString == null)
            {
               throw new Exception("The connection string cannot be null");
            }

            return connectionString.ConnectionString;
        }

        #endregion
    }
}