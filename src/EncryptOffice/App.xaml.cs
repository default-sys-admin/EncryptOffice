using System.Windows;
using DevExpress.Xpf.Core;

namespace EncryptOffice
{
    public partial class App : Application
    {
        private void App_OnStartup(object sender, StartupEventArgs e)
        {
            ApplicationThemeHelper.ApplicationThemeName = Theme.Office2019WhiteName;
        }
    }
}