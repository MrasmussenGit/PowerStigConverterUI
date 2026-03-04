using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Media.Imaging;

namespace PowerStigConverterUI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Set icon for all windows globally
            EventManager.RegisterClassHandler(typeof(Window), Window.LoadedEvent, new RoutedEventHandler(OnWindowLoaded));
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            if (sender is Window window && window.Icon == null)
            {
                window.Icon = new BitmapImage(new Uri("pack://application:,,,/Images/App.ico"));
            }
        }
    }

}
