using System;
using System.Threading.Tasks;
using System.Windows;

namespace Lieferliste_WPF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public partial class App : Application
    {
        public static double GlobalFontSize
        {
            get => (double)Current.Resources["GlobalFontSize"];
            set => Current.Resources["GlobalFontSize"] = value;
        }
        protected override void OnStartup(StartupEventArgs e)
        {
            //ThemeManager.Current.AddLibraryTheme(new LibraryTheme(
            //    new Uri("pack://application:,,,/MahAppsMetroThemesSample;component/CustomAccents/Light.Accent1.xaml"),
            //    MahAppsLibraryThemeProvider.DefaultInstance));
            //ThemeManager.Current.AddLibraryTheme(new LibraryTheme(
            //    new Uri("pack://application:,,,/MahAppsMetroThemesSample;component/CustomAccents/Dark.Accent1.xaml"),
            //    MahAppsLibraryThemeProvider.DefaultInstance));
            //ThemeManager.Current.AddLibraryTheme(new LibraryTheme(
            //    new Uri("pack://application:,,,/MahAppsMetroThemesSample;component/CustomAccents/Light.Accent2.xaml"),
            //    MahAppsLibraryThemeProvider.DefaultInstance));

            base.OnStartup(e);

            //var theme = ThemeManager.Current.AddLibraryTheme(new LibraryTheme(
            //    new Uri("pack://application:,,,/MahAppsMetroThemesSample;component/CustomAccents/Dark.Accent2.xaml"),
            //    MahAppsLibraryThemeProvider.DefaultInstance));
            //var theme = ThemeManager.Current.AddLibraryTheme(
            //new LibraryTheme(
            //    new Uri("pack://application:,,,/Lieferliste_WPF;component/Themes/Light.Accent2.xaml"),
            //    MahAppsLibraryThemeProvider.DefaultInstance
            //    )
            //);
            // Create runtime themes


            //ThemeManager.Current.AddTheme(RuntimeThemeGenerator.Current.GenerateRuntimeTheme("Dark", Colors.Red));
            //ThemeManager.Current.AddTheme(RuntimeThemeGenerator.Current.GenerateRuntimeTheme("Light", Colors.Red));

            //ThemeManager.Current.AddTheme(RuntimeThemeGenerator.Current.GenerateRuntimeTheme("Dark", Colors.GreenYellow));
            //ThemeManager.Current.AddTheme(RuntimeThemeGenerator.Current.GenerateRuntimeTheme("Light", Colors.GreenYellow));

            //ThemeManager.Current.AddTheme(RuntimeThemeGenerator.Current.GenerateRuntimeTheme("Dark", Colors.Indigo));
            //ThemeManager.Current.ChangeTheme(this, ThemeManager.Current.AddTheme(RuntimeThemeGenerator.Current.GenerateRuntimeTheme("Light", Colors.Indigo)));

            //ThemeManager.Current.AddTheme(new Theme("LightOlive", "LightOlive", "Light", "Olive", Colors.Olive, Brushes.Olive, true, false));
            //ThemeManager.Current.AddTheme(new Theme("LightBlue", "LightBlue", "Light", "Blue", Colors.Blue, Brushes.Blue, true, false));
            //ThemeManager.Current.AddTheme(new Theme("LightOrange", "LightOrange", "Light", "Orange", Colors.Orange, Brushes.Orange, true, false));
            //ThemeManager.Current.AddTheme(new Theme("LightSeaGreen", "LightSeaGreen", "Light", "SeaGreen", Colors.SeaGreen, Brushes.SeaGreen, true, false));

            //ThemeManager.Current.AddTheme(new Theme("DarkOlive", "DarkOlive", "Dark", "Olive", Colors.Olive, Brushes.Olive, true, false));
            //ThemeManager.Current.AddTheme(new Theme("DarkBlue", "DarkBlue", "Dark", "Blue", Colors.Blue, Brushes.Blue, true, false));
            //ThemeManager.Current.AddTheme(new Theme("DarkOrange", "DarkOrange", "Dark", "Orange", Colors.Orange, Brushes.Orange, true, false));
            //ThemeManager.Current.AddTheme(new Theme("DarkSeaGreen", "DarkSeaGreen", "Dark", "SeaGreen", Colors.SeaGreen, Brushes.SeaGreen, true, false));

            try
            {
                var bootstrapper = new Bootstrapper();
                bootstrapper.Run();
            }
            catch (Exception ex)
            {
                if (MessageBox.Show(ex.Message,"ERROR", MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK)
                
                    throw;
            }
        }

    }

}

