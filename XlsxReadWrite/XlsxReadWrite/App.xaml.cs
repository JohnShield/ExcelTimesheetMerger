/* Author: John Shield
 * Date: 2014 October
 */

using System.Windows;
using System.Windows.Threading;
using XlsxReadWrite.Properties;

namespace XlsxReadWrite
{
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += this.OnDispatcherUnhandledException;
        }

        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), "Unhandled exception occured", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.None);
            e.Handled = true;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Settings.Default.Save();
            base.OnExit(e);
        }
    }
}
