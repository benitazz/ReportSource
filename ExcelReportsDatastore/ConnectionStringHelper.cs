#region

using System;
using System.Configuration;
using System.Data.EntityClient;
using System.Data.Objects;
using System.Linq;

#endregion

namespace ExcelReportsDatastore
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
            ConfigurationManager.RefreshSection("ConnectionStrings");
            ConnectionStringSettingsCollection settings = ConfigurationManager.ConnectionStrings;

            if (settings != null)
            {
                foreach (ConnectionStringSettings cs in settings)
                {
                    Console.WriteLine(cs.Name);
                    Console.WriteLine(cs.ProviderName);
                    Console.WriteLine(cs.ConnectionString);
                }
            }

            var t = System.Configuration.ConfigurationManager.AppSettings["ExcelReportsDatabaseEntities"];
            var connectionString = ConfigurationManager.ConnectionStrings["Context"];

            if (connectionString == null)
            {
                throw new Exception("The connection string cannot be null");
            }

            // return connectionString.ConnectionString;
            return null;
        }

        /// <summary>
        /// The get connection string.
        /// </summary>
        /// <param name="context">
        /// The context.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public static string GetConnectionString(ObjectContext context)
        {
          /*var t = databaseEntities.ExcelDataTables.FirstOrDefault();
          t.*/
          
          return ((EntityConnection)context.Connection).StoreConnection.ConnectionString;
        }

        #endregion
    }
}