using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;


namespace El2Core.Utils
{
    public abstract class AttachmentFactory
    {
        // SHGetFileInfo flags
        [Flags]
        private enum SHGFI : uint
        {
            Icon = 0x000000100,
            LargeIcon = 0x000000000,
            SmallIcon = 0x000000001,
            UseFileAttributes = 0x000000010
        }

        // File attribute constants
        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        [StructLayout(LayoutKind.Sequential)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SHGetFileInfo(
            string pszPath,
            uint dwFileAttributes,
            ref SHFILEINFO psfi,
            uint cbFileInfo,
            uint uFlags
        );

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);
        public abstract IDisplayAttachment CreateDisplayAttachment(string link, bool isLink);
        public abstract IDbAttachment CreateDbAttachment(string link, bool isLink);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto, PreserveSig = true, SetLastError = false)]
        public static extern IntPtr GetActiveWindow();

        [ComImport]
        [Guid("3E68D4BD-7135-4D10-8018-9FB6D9F33FA1")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IInitializeWithWindow
        {
            void Initialize(IntPtr hwnd);
        }

 

        /// <summary>
        /// Gets the icon for a given file extension.
        /// </summary>
        /// <param name="extension">File extension (e.g., ".txt")</param>
        /// <param name="largeIcon">True for large icon, false for small</param>
        /// <returns>ImageSource for WPF</returns>
        public static ImageSource GetFileTypeIcon(string extension, bool largeIcon)
        {
            if (string.IsNullOrWhiteSpace(extension))
                throw new ArgumentException("Extension cannot be null or empty.", nameof(extension));

            if (!extension.StartsWith("."))
                extension = "." + extension;

            SHFILEINFO shinfo = new SHFILEINFO();
            uint flags = (uint)(SHGFI.Icon | SHGFI.UseFileAttributes |
                                (largeIcon ? SHGFI.LargeIcon : SHGFI.SmallIcon));

            IntPtr hImg = SHGetFileInfo(extension, FILE_ATTRIBUTE_NORMAL, ref shinfo,
                                        (uint)Marshal.SizeOf(shinfo), flags);

            if (hImg == IntPtr.Zero)
                return null;

            try
            {
                Icon icon = Icon.FromHandle(shinfo.hIcon);
                ImageSource img = Imaging.CreateBitmapSourceFromHIcon(
                    icon.Handle,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
                return img;
            }
            finally
            {
                DestroyIcon(shinfo.hIcon); // Prevent memory leak
            }
        }

    
        public static IDisplayAttachment FloatAttachment(IDisplayAttachment attachment, string? file, bool isLink)
        {
            FileInfo fi = new FileInfo(file ?? string.Empty);
 
            attachment.Content = GetFileTypeIcon(fi.Extension, true);
  
            attachment.Name = (isLink) ? fi.FullName : fi.Name;
            attachment.IsLink = isLink;
            
            return attachment; 
        }
        public static IDbAttachment FloatAttachment(IDbAttachment dbAttachment, string fileString, bool isLink)
        {
            FileInfo fi = new FileInfo(fileString);
            if (fi.Exists)
            {
                if (isLink)
                {
                    dbAttachment.Link = fi.FullName;
                    dbAttachment.IsLink = true;
                    dbAttachment.TimeStamp = DateTime.Now;
                }
                else
                {
                    if (fi.Length < 0x500000)    //Filesize of 5 MiB
                    {

                        MemoryStream ms = new MemoryStream();
                        using (FileStream file = new FileStream(fileString, FileMode.Open, FileAccess.Read))
                            file.CopyTo(ms);
                        dbAttachment.Link = fi.Name;
                        dbAttachment.IsLink= false;
                        dbAttachment.BinaryData = ms.ToArray();
                        dbAttachment.TimeStamp = DateTime.Now;
                    }
                    else if (MessageBox.Show("Die Datei ist größer als 5 MiB, soll es als Link gespeichert werden?", "",
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        dbAttachment.Link = fi.FullName;
                        dbAttachment.IsLink = true;
                        dbAttachment.TimeStamp = DateTime.Now;
                    }
                }
            }
            else MessageBox.Show("Datei wurde nicht gefunden", "Datei anfügen", MessageBoxButton.OK, MessageBoxImage.Error);
            return dbAttachment;
        }
        public static void OpenFile(string file, MemoryStream? memoryStream)
        {
            try
            {
                FileInfo fi = new FileInfo(file);
                string filepath;
                if (memoryStream == null)  
                {
                    filepath = fi.FullName;
                }
                else
                {

                    filepath = Path.Combine(Path.GetTempPath(), fi.Name);
                    using FileStream fs = new(filepath, FileMode.Create);
                    memoryStream.CopyTo(fs);
                    fs.Flush();
                    fs.Close();
                }
         
                 new Process() { StartInfo = new ProcessStartInfo(filepath) { UseShellExecute = true } }.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("{0}\n{1}", e.Message, e.InnerException), "OpenStream", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public static async Task<string> GetFilePickerPath()
        {

            FileOpenPicker openPicker = new();

            WinRT.Interop.InitializeWithWindow.
                Initialize(openPicker, new System.Windows.Interop.WindowInteropHelper(Application.Current.MainWindow).Handle);
 
            openPicker.ViewMode = PickerViewMode.List;
            openPicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            openPicker.FileTypeFilter.Add("*");

            StorageFile op = await openPicker.PickSingleFileAsync();
            if (op != null) { return op.Path; }
            

            return string.Empty;
        }


    }
    public interface IDisplayAttachment
    {
        int Id { get; set; }
        string Name { get; set; }
        string? Description { get; set; }
        object? Content { get; set; }
        bool IsLink { get; set; }
    }
    public interface IDbAttachment
    {
        string Link { get; set; }
        bool IsLink { get; set; }
        DateTime TimeStamp { get; set; }
        byte[]? BinaryData { get; set; }
    }

}
