using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using System;
using System.IO;
using System.Windows;

namespace ModuleShift.Services
{
    public static class MonacoHelper
    {
        // 1. Registrierung des Attached Property für das Binding im XAML
        public static readonly DependencyProperty BindableCodeProperty =
            DependencyProperty.RegisterAttached(
                "BindableCode",
                typeof(string),
                typeof(MonacoHelper),
                new FrameworkPropertyMetadata(default(string), FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnBindableCodeChanged));

        public static string GetBindableCode(DependencyObject d) => (string)d.GetValue(BindableCodeProperty);
        public static void SetBindableCode(DependencyObject d, string value) => d.SetValue(BindableCodeProperty, value);

        // Interner Schlüssel, um zu verhindern, dass Updates im Kreis gesendet werden
        private static readonly DependencyProperty IsUpdatingProperty =
            DependencyProperty.RegisterAttached("IsUpdating", typeof(bool), typeof(MonacoHelper), new PropertyMetadata(false));

        // 2. Event-Handler: Wird aufgerufen, wenn sich der Code im C#-ViewModel ändert (C# -> JS)
        private static async void OnBindableCodeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is WebView2 webView)
            {
                // Wenn die Engine noch nicht bereit ist, brechen wir ab. 
                // Der Wert wird stattdessen geladen, sobald "NavigationCompleted" feuert.
                if (webView.CoreWebView2 == null) return;

                bool isUpdating = (bool)webView.GetValue(IsUpdatingProperty);
                if (isUpdating) return;

                string newCode = e.NewValue as string ?? string.Empty;
                string escapedCode = System.Web.HttpUtility.JavaScriptStringEncode(newCode);

                webView.SetValue(IsUpdatingProperty, true);
                await webView.ExecuteScriptAsync($"setCode('{escapedCode}');");
                webView.SetValue(IsUpdatingProperty, false);
            }
        }

        // 3. Initialisierungsmethode für das WebView2-Control (ClickOnce-sicher)
        public static readonly DependencyProperty InitializeWithCodeProperty =
            DependencyProperty.RegisterAttached(
                "InitializeWithCode",
                typeof(bool),
                typeof(MonacoHelper),
                new PropertyMetadata(false, OnInitializeWithCodeChanged));

        public static bool GetInitializeWithCode(DependencyObject d) => (bool)d.GetValue(InitializeWithCodeProperty);
        public static void SetInitializeWithCode(DependencyObject d, bool value) => d.SetValue(InitializeWithCodeProperty, value);

        private static async void OnInitializeWithCodeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is WebView2 webView && (bool)e.NewValue)
            {
                // ClickOnce-sicheres Caching im Benutzerverzeichnis
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string cacheFolder = Path.Combine(localAppData, "WpfMonacoApp_WebView2_Cache");

                try
                {
                    var env = await CoreWebView2Environment.CreateAsync(userDataFolder: cacheFolder);
                    await webView.EnsureCoreWebView2Async(env);

                    // Lokales Dateimapping für ClickOnce
                    string assemblyFolder = AppDomain.CurrentDomain.BaseDirectory;
                    webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "monaco.local",
                        assemblyFolder,
                        CoreWebView2HostResourceAccessKind.Allow
                    );
                    // HTML laden
                    webView.CoreWebView2.Navigate("http://monaco.local/index.html");

                    // Event-Handler: Erst wenn die Seite im Browser komplett geladen ist, 
                    // übertragen wir den initialen Code aus dem ViewModel in den Editor.
                    webView.NavigationCompleted += (s, ev) =>
                    {
                        string currentCodeInViewModel = GetBindableCode(webView);
                        if (!string.IsNullOrEmpty(currentCodeInViewModel))
                        {
                            // Rufe die JS-Funktion auf, da CoreWebView2 nun garantiert bereit ist
                            string escapedCode = System.Web.HttpUtility.JavaScriptStringEncode(currentCodeInViewModel);
                            webView.ExecuteScriptAsync($"setCode('{escapedCode}');");
                        }
                    };

                    // Event-Handler: Wenn der Benutzer tippt (JS -> C#)
                    webView.CoreWebView2.WebMessageReceived += (s, args) =>
                    {
                        webView.SetValue(IsUpdatingProperty, true);
                        string codeFromEditor = args.TryGetWebMessageAsString();
                        SetBindableCode(webView, codeFromEditor);
                        webView.SetValue(IsUpdatingProperty, false);
                    };
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler beim Initialisieren des Monaco Editors: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
