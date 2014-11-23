#region

using System;
using System.Data;
using System.Data.SqlServerCe;

#endregion

namespace ExcelReportsDatastore
{
    /// <summary>
    /// The database reader.
    /// </summary>
    public class DatabaseReader
    {
        #region Public Methods and Operators

        /// <summary>
        /// The get database data table.
        /// </summary>
        /// <param name="query">
        /// The query.
        /// </param>
        /// <returns>
        /// The <see cref="DataTable"/>.
        /// </returns>
        public static DataTable GetDatabaseDataTable(string query)
        {
            var connectionString = string.Empty;
            try
            {
                using (var con = new SqlCeConnection(connectionString))
                {
                    using (var cmd = new SqlCeCommand(query))
                    {
                        using (var sda = new SqlCeDataAdapter())
                        {
                            cmd.Connection = con;
                            sda.SelectCommand = cmd;
                            using (var dataTable = new DataTable())
                            {
                                sda.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
        }

        /// <summary>
        /// Gets the database data table.
        /// </summary>
        /// <param name="connection">The connection.</param>
        /// <param name="query">The query.</param>
        /// <returns>
        /// Returns the data table from database query.
        /// </returns>
       public static DataTable GetDatabaseDataTable(SqlCeConnection connection, string query)
        {
            try
            {
                using (var cmd = new SqlCeCommand(query))
                {
                    using (var sda = new SqlCeDataAdapter())
                    {
                        cmd.Connection = connection;
                        sda.SelectCommand = cmd;
                        using (var dataTable = new DataTable())
                        {
                            sda.Fill(dataTable);
                            return dataTable;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
        }

        #endregion
    }
}