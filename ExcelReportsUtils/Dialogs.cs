#region

using System;
using System.Windows.Forms;

using Microsoft.SqlServer.MessageBox;

#endregion

namespace ExcelReportsUtils
{
    /// <summary>
    /// The dialogs.
    /// </summary>
    public static class Dialogs
    {
        #region Constants

        /// <summary>
        /// The message header.
        /// </summary>
        public const string MessageHeader = "Excel Report Manager";

        #endregion

        #region Enums

        /// <summary>
        /// The notification level.
        /// </summary>
        public enum NotificationLevel
        {
            /// <summary>
            /// The user.
            /// </summary>
            User, 

            /// <summary>
            /// The information.
            /// </summary>
            Information, 

            /// <summary>
            /// The warning.
            /// </summary>
            Warning, 

            /// <summary>
            /// The error.
            /// </summary>
            Error
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void ShowError(string message)
        {
            ShowError(message, MessageHeader);
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowError(string message, string caption)
        {
            ShowError(null, message, caption);
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowError(IWin32Window owner, string message, string caption)
        {
            try
            {
                MessageBox.Show(owner, message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                // display message without owner
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="notificationLevel">
        /// The notification level.
        /// </param>
        public static void ShowError(string message, NotificationLevel notificationLevel)
        {
            ShowError(null, message, notificationLevel);
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="notificationLevel">
        /// The notification level.
        /// </param>
        public static void ShowError(IWin32Window owner, string message, NotificationLevel notificationLevel)
        {
            try
            {
                MessageBox.Show(
                    owner, 
                    message, 
                    string.Format("{0} Notification", notificationLevel), 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
            catch
            {
                // display message without owner
                MessageBox.Show(
                    message, 
                    string.Format("{0} Notification", notificationLevel), 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="exception">
        /// The exception.
        /// </param>
        public static void ShowError(Exception exception)
        {
            var mbox = new ExceptionMessageBox(exception);
            mbox.Show(null);

            //  ShowError(ex, MessageHeader);
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="ex">
        /// The ex.
        /// </param>
        public static void ShowError(IWin32Window owner, Exception ex)
        {
            ShowError(owner, ex, MessageHeader);
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="exception">
        /// The ex.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowError(Exception exception, string caption)
        {
            // ShowError(null, ex, caption);

            var mbox = new ExceptionMessageBox(exception) { Caption = caption };
            mbox.Show(null);

            // ExceptionDialog.ShowMessage(caption, ex.Message, ex);
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="ex">
        /// The ex.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowError(IWin32Window owner, Exception ex, string caption)
        {
            try
            {
                MessageBox.Show(owner, ex.Message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

                // ExceptionDialog.ShowMessage(owner, caption, ex.Message, ex);
            }
            catch
            {
                MessageBox.Show(owner, ex.Message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                // display message without owner
                // ExceptionDialog.ShowMessage(caption, ex.Message, ex);
            }
        }

        /// <summary>
        /// The show error.
        /// </summary>
        /// <param name="ex">
        /// The ex.
        /// </param>
        /// <param name="notificationLevel">
        /// The notification level.
        /// </param>
        public static void ShowError(Exception ex, NotificationLevel notificationLevel)
        {
            /*// Define a new top-level error message.
            string str = "Write the reason why the action failed.";

            // Add the new top-level message to the handled exception.
            ApplicationException exTop = new ApplicationException(str, ex) { Source = this.Text };

            // Information in the Data property of an exception that has a name
            // beginning with "HelpLink.Advanced" is shown when the user
            // clicks the Advanced Information button of the exception message
            // box dialog box.
            exTop.Data.Add("AdvancedInformation.FileName",
                "ExceptionMessageBoxSample.dll");
            exTop.Data.Add("AdvancedInformation.FilePosition", "line 24");
            exTop.Data.Add("AdvancedInformation.UserContext",
                "a detail message can be given");

            // Show the exception message box with additional information that
            // is helpful when a user calls technical support.
            ExceptionMessageBox box = new ExceptionMessageBox(exTop)
                                          {
                                              Buttons = ExceptionMessageBoxButtons.YesNo,
                                              Caption = "Caption",
                                              ShowCheckBox = true,
                                              ShowToolBar = true,
                                              Symbol = ExceptionMessageBoxSymbol.Stop
                                          };
            box.Show();

            // ExceptionDialog.ShowMessage(string.Format("{0} Notification", notificationLevel.ToString()), ex.Message, ex);*/
        }

        /// <summary>
        /// The show information.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void ShowInformation(string message)
        {
            ShowInformation(null, message, MessageHeader);
        }

        /// <summary>
        /// The show information.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowInformation(string message, string caption)
        {
            ShowInformation(null, message, caption);
        }

        /// <summary>
        /// The show information.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowInformation(IWin32Window owner, string message, string caption)
        {
            try
            {
                // CMStatic.CloseDialogue();
                MessageBox.Show(owner, message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                // display message without owner
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// The show information.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="notificationLevel">
        /// The notification level.
        /// </param>
        public static void ShowInformation(string message, NotificationLevel notificationLevel)
        {
            ShowInformation(null, message, string.Format("{0} Notification", notificationLevel.ToString()));
        }

        /// <summary>
        /// The show ok cancel.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowOkCancel(string message)
        {
            return ShowOkCancel(message, MessageHeader);
        }

        /// <summary>
        /// The show ok cancel.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowOkCancel(string message, string caption)
        {
            return MessageBox.Show(null, message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
        }

        /// <summary>
        /// The show warning.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        public static void ShowWarning(string message)
        {
            ShowWarning(null, message, MessageHeader);
        }

        /// <summary>
        /// The show warning.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowWarning(string message, string caption)
        {
            ShowWarning(null, message, caption);
        }

        /// <summary>
        /// The show warning.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        public static void ShowWarning(IWin32Window owner, string message, string caption)
        {
            try
            {
                // CMStatic.CloseDialogue();
                MessageBox.Show(owner, message, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch
            {
                // display message without owner
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// The show warning.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="notificationLevel">
        /// The notification level.
        /// </param>
        public static void ShowWarning(string message, NotificationLevel notificationLevel)
        {
            ShowWarning(null, message, string.Format("{0} Notification", notificationLevel.ToString()));
        }

        /// <summary>
        /// The show yes no.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowYesNo(string message, string caption)
        {
            return ShowYesNo(null, message, caption);
        }

        /// <summary>
        /// The show yes no.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowYesNo(IWin32Window owner, string message, string caption)
        {
            try
            {
                // CMStatic.CloseDialogue();
                return MessageBox.Show(owner, message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            catch
            {
                // display message without owner
                return MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
        }

        /// <summary>
        /// The show yes no.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowYesNo(string message)
        {
            return ShowYesNo(message, MessageHeader);
        }

        /// <summary>
        /// The show yes no.
        /// </summary>
        /// <param name="owner">
        /// The owner.
        /// </param>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowYesNo(IWin32Window owner, string message)
        {
            return ShowYesNo(owner, message, MessageHeader);
        }

        /// <summary>
        /// The show yes no cancel.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowYesNoCancel(string message)
        {
            return ShowYesNoCancel(message, MessageHeader);
        }

        /// <summary>
        /// The show yes no cancel.
        /// </summary>
        /// <param name="message">
        /// The message.
        /// </param>
        /// <param name="caption">
        /// The caption.
        /// </param>
        /// <returns>
        /// The <see cref="DialogResult"/>.
        /// </returns>
        public static DialogResult ShowYesNoCancel(string message, string caption)
        {
            return MessageBox.Show(null, message, caption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        }

        #endregion
    }
}