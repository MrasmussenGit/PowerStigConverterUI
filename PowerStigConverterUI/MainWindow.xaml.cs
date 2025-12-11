using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerStigConverterUI
{
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();

            var entry = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
            var ver = entry.GetName().Version?.ToString() ?? "unknown";
            VersionTextBlock.Text = $"Version {ver}";
        }

        private void ConvertStigButton_Click(object sender, RoutedEventArgs e)
        {
            var win = new ConvertStigWindow { Owner = this };
            win.ShowDialog();
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            var win = new CompareWindow { Owner = this };
            win.ShowDialog();
            //System.Windows.MessageBox.Show("Compare clicked.", "PowerStig Converter", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}