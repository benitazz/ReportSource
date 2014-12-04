#region

using System;
using System.Configuration;
using System.Data.EntityClient;
using System.Data.Objects;
using System.Data.SqlServerCe;

#endregion

namespace ExcelReportsDatastore.Helpers
{
  /// <summary>
  /// The connection string helper.
  /// </summary>
  public class ConnectionStringHelper
  {
    #region Public Properties

    /// <summary>
    /// Gets the excel database connection string.
    /// </summary>
    /// <value>
    /// The excel database connection string.
    /// </value>
    /// <exception cref="System.Exception">The connection string cannot be null</exception>
    public static string ExcelDatabaseConnectionString
    {
      get
      {
        var connectionString = ConfigurationManager.ConnectionStrings["ExcelDatabaseConnectionString"];

        if (connectionString == null)
        {
          throw new Exception("The connection string cannot be null");
        }

        var engine =
         new SqlCeEngine(connectionString.ConnectionString);

        if (engine.Verify())
        {
          return connectionString.ConnectionString;
        }

        Console.WriteLine("Database is corrupted.");
        try
        {
          engine.Repair(null, RepairOption.DeleteCorruptedRows);
        }
        catch (SqlCeException ex)
        {
          Console.WriteLine(ex.Message);
        }

        return connectionString.ConnectionString;
      }
    }

    #endregion

    #region Public Methods and Operators

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

    #region Methods

    /// <summary>
    /// The get connection string.
    /// </summary>
    /// <returns>
    /// The <see cref="string"/>.
    /// </returns>
    private static string GetConnectionString()
    {
      if (!string.IsNullOrEmpty(ExcelDatabaseConnectionString))
      {
        return ExcelDatabaseConnectionString;
      }

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

      /*var appConfig = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
      string dllConfigData = appConfig.AppSettings.Settings["ExcelReportsDatastore"].Value;

      var asmName = System.Reflection.Assembly.GetAssembly(typeof(ConnectionStringHelper)).GetName().Name;
      var asmPath = System.Web.HttpContext.Current.Server.MapPath(@"bin\" + asmName + ".dll");
      var cm = ConfigurationManager.OpenExeConfiguration(asmPath);
      var test = cm.AppSettings.Settings["Context"].Value;

      var t = ConfigurationManager.AppSettings["ExcelReportsDatabaseEntities"];*/
      var connectionString = ConfigurationManager.ConnectionStrings["Context"];

      if (connectionString == null)
      {
        throw new Exception("The connection string cannot be null");
      }

      return connectionString.ConnectionString;
    }

    #endregion
  }
}