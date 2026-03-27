using System;
using System.Threading.Tasks;
using System.Windows;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop; // Required for the interop helpers

namespace El2Core.Services
{
    public interface IFilePickerService
    {
        Task<string?> PickFileAsync();
    }

    public class FilePickerService : IFilePickerService
    {


        public async Task<string?> PickFileAsync()
        {
            FileOpenPicker openPicker = new();

            // 1. Get the window handle (HWND) from the current active window.
            // This is the modern replacement for the WPF-specific WindowInteropHelper.
            var window = Application.Current.MainWindow; // A helper method to get the window
            var hwnd = WindowNative.GetWindowHandle(window);

            // 2. Initialize the file picker with the window handle.
            InitializeWithWindow.Initialize(openPicker, hwnd);

            // 3. Configure and show the picker (this part is the same).
            openPicker.ViewMode = PickerViewMode.List;
            openPicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            openPicker.FileTypeFilter.Add("*");

            StorageFile file = await openPicker.PickSingleFileAsync();

            // 4. Return the path, or null if the user cancelled.
            return file?.Path;
        }
    }
}
