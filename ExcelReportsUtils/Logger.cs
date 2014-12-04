#region

using System;
using System.Collections.Generic;
using System.Globalization;

using log4net;
using log4net.Config;

#endregion

namespace ExcelReportsUtils
{
    /// <summary>
    /// The logger.
    /// </summary>
    public class Logger
    {
        #region Static Fields

        /// <summary>
        /// The _lock.
        /// </summary>
        private static readonly object _lock = new object();

        /// <summary>
        /// The _loggers.
        /// </summary>
        private static readonly Dictionary<Type, ILog> _loggers = new Dictionary<Type, ILog>();

        /// <summary>
        /// The _log initialized.
        /// </summary>
        private static bool _logInitialized;

        #endregion

        /* Log a message object */
        #region Public Methods and Operators

        /// <summary>
        /// The debug.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Debug(object source, string message)
        {
            Debug(source.GetType(), message);
        }

        /// <summary>
        /// The debug.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="ps">
        /// The ps.
        /// </param>
        public static void Debug(object source, string message, params object[] ps)
        {
            Debug(source.GetType(), string.Format(message, ps));
        }

        /// <summary>
        /// The debug.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Debug(Type source, string message)
        {
            ILog logger = getLogger(source);

            if (logger.IsDebugEnabled)
            {
                logger.Debug(message);
            }
        }

        /// <summary>
        /// The debug.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Debug(object source, object message, Exception exception)
        {
            Debug(source.GetType(), message, exception);
        }

        /// <summary>
        /// The debug.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Debug(Type source, object message, Exception exception)
        {
            getLogger(source).Debug(message, exception);
        }

        /// <summary>
        /// The ensure initialized.
        /// </summary>
        public static void EnsureInitialized()
        {
            if (!_logInitialized)
            {
                initialize();
            }
        }

        /// <summary>
        /// The error.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Error(object source, object message)
        {
            Error(source.GetType(), message);
        }

        /// <summary>
        /// The error.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Error(Type source, object message)
        {
            ILog logger = getLogger(source);

            if (logger.IsErrorEnabled)
            {
                logger.Error(message);
            }
        }

        /// <summary>
        /// The error.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Error(object source, object message, Exception exception)
        {
            Error(source.GetType(), message, exception);
        }

        /// <summary>
        /// The error.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Error(Type source, object message, Exception exception)
        {
            getLogger(source).Error(message, exception);
        }

        /// <summary>
        /// The fatal.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Fatal(object source, object message)
        {
            Fatal(source.GetType(), message);
        }

        /// <summary>
        /// The fatal.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Fatal(Type source, object message)
        {
            ILog logger = getLogger(source);

            if (logger.IsFatalEnabled)
            {
                logger.Fatal(message);
            }
        }

        /// <summary>
        /// The fatal.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Fatal(object source, object message, Exception exception)
        {
            Fatal(source.GetType(), message, exception);
        }

        /// <summary>
        /// The fatal.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Fatal(Type source, object message, Exception exception)
        {
            getLogger(source).Fatal(message, exception);
        }

        /// <summary>
        /// The info.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Info(object source, object message)
        {
            Info(source.GetType(), message);
        }

        /// <summary>
        /// The info.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Info(Type source, object message)
        {
            ILog logger = getLogger(source);

            if (logger.IsInfoEnabled)
            {
                logger.Info(message);
            }
        }

        /// <summary>
        /// The info.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Info(object source, object message, Exception exception)
        {
            Info(source.GetType(), message, exception);
        }

        /// <summary>
        /// The info.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Info(Type source, object message, Exception exception)
        {
            getLogger(source).Info(message, exception);
        }

        /// <summary>
        /// The serialize exception.
        /// </summary>
        /// <param name="exception">
        /// The exception.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public static string SerializeException(Exception exception)
        {
            return SerializeException(exception, string.Empty);
        }

        /// <summary>
        /// The warn.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Warn(object source, object message)
        {
            Warn(source.GetType(), message);
        }

        /// <summary>
        /// The warn.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void Warn(Type source, object message)
        {
            ILog logger = getLogger(source);
            if (logger.IsWarnEnabled)
            {
                logger.Warn(message);
            }
        }

        /// <summary>
        /// The warn.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Warn(object source, object message, Exception exception)
        {
            Warn(source.GetType(), message, exception);
        }

        /// <summary>
        /// The warn.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void Warn(Type source, object message, Exception exception)
        {
            getLogger(source).Warn(message, exception);
        }

        #endregion

        #region Methods

        /// <summary>
        /// The serialize exception.
        /// </summary>
        /// <param name="e">
        /// The e.
        /// </param>
        /// <param name="exceptionMessage">
        /// The exception message.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private static string SerializeException(Exception e, string exceptionMessage)
        {
            if (e == null)
            {
                return string.Empty;
            }

            exceptionMessage = string.Format(
                CultureInfo.InvariantCulture, 
                "{0}{1}{2}\n{3}", 
                exceptionMessage, 
                string.IsNullOrEmpty(exceptionMessage) ? string.Empty : "\n\n", 
                e.Message, 
                e.StackTrace);

            if (e.InnerException != null)
            {
                exceptionMessage = SerializeException(e.InnerException, exceptionMessage);
            }

            return exceptionMessage;
        }

        /// <summary>
        /// The get logger.
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <returns>
        /// The <see cref="ILog"/>.
        /// </returns>
        private static ILog getLogger(Type source)
        {
            /*initialize();
            return log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);*/
            lock (_lock)
            {
                if (_loggers.ContainsKey(source))
                {
                    return _loggers[source];
                }

                ILog logger = LogManager.GetLogger(source);
                _loggers.Add(source, logger);

                initialize();
                return logger;
            }
        }

        /// <summary>
        /// The initialize.
        /// </summary>
        private static void initialize()
        {
            /*log4net.Config.XmlConfigurator.ConfigureAndWatch(new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "logging.config"));*/
            XmlConfigurator.Configure();

            /*XmlConfigurator.ConfigureAndWatch(
                new FileInfo(Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "Log4Net.config")));*/
            _logInitialized = true;
        }

        #endregion
    }
}