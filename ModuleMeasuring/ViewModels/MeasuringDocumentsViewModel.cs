using DocumentFormat.OpenXml.Packaging;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using GongSolutions.Wpf.DragDrop;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using OpenXmlPowerTools;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml.Linq;


namespace ModuleMeasuring.ViewModels
{
    class MeasuringDocumentsViewModel : ViewModelBase, IDropTarget, INavigationAware
    {
        public string Title { get; } = "Messdokumente";
        private MeasureFirstPartInfo FirstPartInfo { get; set; }
        private VmpbDocumentInfo VmpbInfo { get; set; }
        private MeasureDocumentInfo MeasureInfo { get; set; }
        public MeasuringDocumentsViewModel(IContainerExtension container)
        {
            _container = container;
            var factory = container.Resolve<ILoggerFactory>();
            _logger = factory.CreateLogger<MeasuringDocumentsViewModel>();

            VmpbCreateCommand = new ActionCommand(onVmpbExecuted, onVmpbCanExecute);
            VmpbDeleteCommand = new ActionCommand(onVmpbDelExecuted, onVmpbDelCanExecute);
            VmpbCreatePdfCommand = new ActionCommand(onVmpbCreatePdfExecuted, onVmpbCreatePdfCanExecute);
            VmpbOriginalOpenCommand = new ActionCommand(onVmpbOpenExecuted, onVmpbOpenCanExecute);
            VmpbCloseCommand = new ActionCommand(onVmpbCloseExecuted, onVmpbCloseCanExecute);
            PruefDataCommand = new ActionCommand(onPruefExecuted, onPruefCanExecute);
            PruefDataVmpbCommand = new ActionCommand(onPruefVmpbExecuted, onPruefVmpbCanExecute);
            OpenFileCommand = new ActionCommand(onOpenFileExecuted, onOpenFileCanExecute);
            DeleteFileCommand = new ActionCommand(onDeleteFileExecuted, onDeleteFileCanExecute);
            AddFileCommand = new ActionCommand(onAddFileExecuted, onAddFileCanExecute);
            OpenZngCommand = new ActionCommand(onOpenZngExecuted, onOpenZngCanExecute);
            ConvertToPdfCommand = new ActionCommand(onConvertToPdfExecuted, onConvertToPdfCanExecute);
            LoadData();
            FirstPartInfo = new MeasureFirstPartInfo(_container);
            VmpbInfo = new VmpbDocumentInfo(_container);
            MeasureInfo = new MeasureDocumentInfo(_container);
            
            _watcherFirst.Filter = "*.*";
            _watcherFirst.NotifyFilter = NotifyFilters.LastWrite;
            _watcherFirst.Changed += OnChanged;
            _watcherVmpb.Filter = "*.*";
            _watcherVmpb.NotifyFilter = NotifyFilters.LastWrite;
            _watcherVmpb.Changed += OnChanged;
            _watcherPart.Filter = "*.*";
            _watcherPart.NotifyFilter = NotifyFilters.LastWrite;
            _watcherPart.Changed += OnChanged;
            _watcherScan.Filter = "*.pdf";
            _watcherScan.NotifyFilter = NotifyFilters.LastWrite;
            _watcherScan.Changed += OnChanged;
        }

        IContainerExtension _container;
        private ILogger _logger;

        public ICommand? VmpbCreateCommand { get; private set; }
        public ICommand? VmpbDeleteCommand { get; private set; }
        public ICommand? VmpbCreatePdfCommand { get; private set; }
        public ICommand? VmpbOriginalOpenCommand { get; private set; }
        public ICommand? VmpbCloseCommand { get; private set; }
        public ICommand? PruefDataCommand { get; private set; }
        public ICommand? PruefDataVmpbCommand { get; private set; }
        public ICommand? OpenFileCommand { get; private set; }
        public ICommand? DeleteFileCommand { get; private set; }
        public ICommand? AddFileCommand { get; private set; }
        public ICommand? OpenZngCommand { get; private set; }
        public ICommand? ConvertToPdfCommand { get; private set; }
        public IEnumerable<dynamic> ProcessList { get; private set; }
        private ICollectionView ProcessListView { get; set; }
        private CollectionViewSource processViewSource { get; } = new();
        private ObservableCollection<DocumentDisplay> _FirstDocumentItems = [];
        public ICollectionView FirstDocumentItems { get; private set; }
        private ObservableCollection<DocumentDisplay> _VmpbDocumentItems = [];
        public ICollectionView VmpbDocumentItems { get; private set; }
        private ObservableCollection<DocumentDisplay> _PartDocumentItems = [];
        public ICollectionView PartDocumentItems { get; private set; }
        private FileSystemWatcher _watcherPart = new();
        private FileSystemWatcher _watcherVmpb = new();
        private FileSystemWatcher _watcherFirst = new();
        private FileSystemWatcher _watcherScan = new();
        private VorgItem _SelectedItem;

        public VorgItem SelectedItem
        {
            get
            {
                return _SelectedItem;
            }
            set
            {
                if (value != _SelectedItem)
                {
                    _SelectedItem = value;
                    InWorkState = value.SourceVorgang?.VorgangDocu?.InWorkState ?? 0;
                    SelectedValue = _SelectedItem.Auftrag;
                    NotifyPropertyChanged(() => SelectedItem);
                    OnOrderChanged();
                }
            }
        }


        private string? _SelectedValue;

        public string? SelectedValue
        {
            get
            {
                return _SelectedValue;
            }
            set
            {
                if (value != _SelectedValue)
                {
                    _SelectedValue = value;
                    NotifyPropertyChanged(() => SelectedValue);
                }
            }
        }
        private int _InworkState = 0;

        public int InWorkState
        {
            get { return _InworkState; }
            set
            {
                if (value != _InworkState)
                {
                    _InworkState = value;
                    NotifyPropertyChanged(() => InWorkState);
                }
            }
        }

        private void LoadData()
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            ProcessList = [.. db.Vorgangs.AsNoTracking()
                .Include(x => x.AidNavigation)
                .ThenInclude(x => x.MaterialNavigation)
                .Include(x => x.AidNavigation.DummyMatNavigation)
                .Include(x => x.VorgangDocu)
                .Where(x => x.AidNavigation.Abgeschlossen == false)
                .OrderBy(x => x.Aid)
                .ThenBy(x => x.Vnr)
                .Select(s => new VorgItem(s))];

            //ProcessListView = CollectionViewSource.GetDefaultView(ProcessList);
            //ProcessListView.CurrentChanged += OnOrderChanged;
            FirstDocumentItems = CollectionViewSource.GetDefaultView(_FirstDocumentItems);
            VmpbDocumentItems = CollectionViewSource.GetDefaultView(_VmpbDocumentItems);
            PartDocumentItems = CollectionViewSource.GetDefaultView(_PartDocumentItems);
            
        }
        private void OnChanged(object sender, FileSystemEventArgs e)
        {

        }
        private bool onOpenZngCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddPruefDoc) &&
                SelectedItem != null;
        }

        private void onOpenZngExecuted(object obj)
        {
            try
            {
                string link = "https://fe0-dbb-p-web.rbesz01.com/webcube/?action=showsearch&id=8b6ce5da-1e74-40dd-acba-0b7120f6845e&reusesession=1&samllogin=1&left=1&documentnumber=";
                link += SelectedItem.Material[..10];
                var pi = new ProcessStartInfo(link)
                {
                    UseShellExecute = true,
                    Verb = "OPEN"
                };
                Process.Start(pi);
                //var docu = FirstPartInfo.CreateDocumentInfos([SelectedItem.AidNavigation.Material, SelectedItem.Aid]);
                //string source = Path.Combine(docu[DocumentPart.RasterFolder1], docu[DocumentPart.SavePath]);
                //source = source.TrimEnd(Path.DirectorySeparatorChar);
                //string target = Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], docu[DocumentPart.TTNR]);
                //int i = 0;
                //do
                //{
                //    var so = source + i.ToString() + ".pdf";
                //    if (File.Exists(so))
                //    {
                //        var ta = target + "-" + i.ToString() + ".pdf";
                //        File.Copy(so, ta, true);
                //        _FirstDocumentItems.Add(new DocumentDisplay() { Display = docu[DocumentPart.TTNR], FullName = ta });
                //    }
                //    else { MessageBox.Show(string.Format("Datei {0} wurde nicht gefunden", source), "Raster Copy", MessageBoxButton.OK); }
                //    i++;
                //} while (File.Exists(source + i.ToString() + ".pdf"));                
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "Raster Copy", MessageBoxButton.OK);
            }
        }
        private bool onConvertToPdfCanExecute(object arg)
        {
            return true;
        }

        private void onConvertToPdfExecuted(object obj)
        {
            // Spire.doc

            //Spire.Doc.Document document = new Spire.Doc.Document();
            var doc = VmpbInfo.GetDocument();
            //document.LoadFromFile(doc[DocumentPart.File]);
            //var targ = doc[DocumentPart.File].Replace(".dotx", ".pdf");
            //document.SaveToFile(targ, Spire.Doc.FileFormat.PDF);


            var source = Package.Open(doc[DocumentPart.File]);
            var document = WordprocessingDocument.Open(source);
            HtmlConverterSettings settings = new HtmlConverterSettings();
            XElement html = HtmlConverter.ConvertToHtml(document, settings);

            Console.WriteLine(html.ToString());
            var writer = File.CreateText(doc[DocumentPart.File].Replace(".docx", ".html"));
            writer.WriteLine(html.ToString());
            writer.Dispose();
            Console.ReadLine();

            //using GroupDocs.Conversion;
            //using GroupDocs.Conversion.Options.Convert;

            //var converter = new Converter(@"C:\Pfad\zu\Dokument.docx");
            //var options = new PdfConvertOptions();
            //converter.Convert(@"C:\Pfad\zu\Dokument.pdf", options);


            //var doc = new HtmlToPdfDocument()
            //{
            //            GlobalSettings = {
            //    ColorMode = ColorMode.Color,
            //    Orientation = Orientation.Landscape,
            //    PaperSize = PaperKind.A4,
            //    },
            //            Objects = {
            //    new ObjectSettings() {
            //        PagesCount = true,
            //        HtmlContent = File.ReadAllText(@"C:\TFS\Sandbox\Open-Xml-PowerTools-abfbaac510d0d60e2f492503c60ef897247716cf\ToolsTest\test1.html"),
            //        WebSettings = { DefaultEncoding = "utf-8" },
            //        HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
            //        FooterSettings = { FontSize = 9, Right = "Page [page] of [toPage]" }
            //    }
            //}
            //};
        }
        private bool onAddFileCanExecute(object arg)
        {
            var target = arg as string;
            bool accept;
            switch (target)
            {
                case "first":
                    accept = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddPruefDoc) &&
                        SelectedItem != null; break;
                case "vmpb":
                    accept = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddVmpb) &&
                        SelectedItem != null; break;
                case "part":
                    accept = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddMeasureDocu) &&
                        SelectedItem != null; break;
                default: accept = false; break;
            }
            return accept;
        }

        private void onAddFileExecuted(object obj)
        {
            var target = obj as string;
            if (target != null)
            {
                string jump;
                var dialog = new Microsoft.Win32.OpenFileDialog();
                var setting = new UserSettingsService();
                switch (target)
                {
                    case "first":
                        var Fdocu = FirstPartInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                        FirstPartInfo.Collect();
                        if (string.IsNullOrEmpty(setting.PersonalFolder))
                        {
                            switch (Fdocu[DocumentPart.JumpTarget].ToUpperInvariant())
                            {
                                case "DESKTOP":
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); break;
                                case "DOKUMENTE":
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.Personal); break;
                                default:
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); break;
                            }
                        } else { jump = setting.PersonalFolder; }

                        dialog.InitialDirectory = jump;
                        bool? Fresult = dialog.ShowDialog();
                        if (Fresult == true)
                        {
                            FileInfo fileInfo = new FileInfo(dialog.FileName);
                            var FilePath = Path.Combine(Fdocu[DocumentPart.RootPath], Fdocu[DocumentPart.SavePath], fileInfo.Name);
                            File.Copy(dialog.FileName, FilePath);

                            _FirstDocumentItems.Add(new DocumentDisplay(dialog.FileName, fileInfo.Name));
                        }
                    break;
                    case "vmpb":
                        var VMdocu = VmpbInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag, SelectedItem.Vorgang]);
                        VmpbInfo.Collect();
                        if (string.IsNullOrEmpty(setting.PersonalFolder))
                        {
                            switch (VMdocu[DocumentPart.JumpTarget].ToUpperInvariant())
                            {
                                case "DESKTOP":
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); break;
                                case "DOKUMENTE":
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.Personal); break;
                                default:
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); break;
                            }
                        } else { jump = setting.PersonalFolder; }
                            dialog.InitialDirectory = jump;
                        bool? Vresult = dialog.ShowDialog();
                        if (Vresult == true)
                        {
                            FileInfo fileInfo = new FileInfo(dialog.FileName);
                            var vmFilePath = Path.Combine(VMdocu[DocumentPart.RootPath], VMdocu[DocumentPart.SavePath], fileInfo.Name);
                            File.Copy(dialog.FileName, vmFilePath);

                            _VmpbDocumentItems.Add(new DocumentDisplay(dialog.FileName, fileInfo.Name ));
                        }
                    break;
                    case "part":
                        var Mdocu = MeasureInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                        MeasureInfo.Collect();
                        if (string.IsNullOrEmpty(setting.PersonalFolder))
                        {
                            switch (Mdocu[DocumentPart.JumpTarget].ToUpperInvariant())
                            {
                                case "DESKTOP":
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); break;
                                case "DOKUMENTE":
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.Personal); break;
                                default:
                                    jump = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); break;
                            }
                        } else { jump = setting.PersonalFolder; }
                        dialog.InitialDirectory = jump;
                        bool? Mresult = dialog.ShowDialog();
                        if (Mresult == true)
                        {
                            FileInfo fileInfo = new FileInfo(dialog.FileName);
                            var mFilePath = Path.Combine(Mdocu[DocumentPart.RootPath], Mdocu[DocumentPart.SavePath], fileInfo.Name);
                            File.Copy(dialog.FileName, mFilePath);

                            _PartDocumentItems.Add(new DocumentDisplay(dialog.FileName, fileInfo.Name ));
                        }
                        break;
                }
            }
        }

        private bool onDeleteFileCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.DelMeasureDocu);
        }

        private void onDeleteFileExecuted(object obj)
        {
            try
            {
                if(obj is DocumentDisplay dis )
                {
                    File.Delete(dis.FullName);
                    _FirstDocumentItems.Remove(dis);
                    _VmpbDocumentItems.Remove(dis);
                    _PartDocumentItems.Remove(dis);
                }

            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "Delete File", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private bool onOpenFileCanExecute(object arg)
        {
            return true;
        }

        private void onOpenFileExecuted(object obj)
        {
            if (obj is DocumentDisplay s) new Process() { StartInfo = new ProcessStartInfo(s.FullName) { UseShellExecute = true } }.Start();
        }
        private bool onPruefCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddPruefDoc) &&
                SelectedItem != null && _FirstDocumentItems.All(x => x.DocumentType != 1);
        }

        private void onPruefExecuted(object obj)
        {
            var mes = ProcessList.First(x => x.Auftrag == SelectedItem.Auftrag);  
            var docu = FirstPartInfo.CreateDocumentInfos([mes.Material, mes.Auftrag]);
            FirstPartInfo.Collect();
            Directory.CreateDirectory(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], docu[DocumentPart.Folder]));
            FileInfo Firstfile = new FileInfo(docu[DocumentPart.Template]);
            var Firsttarg = new FileInfo(docu[DocumentPart.File]);
            if (!Firsttarg.Exists)
                File.Copy(Firstfile.FullName, Firsttarg.FullName);

            _FirstDocumentItems.Clear();
            foreach (var d in Firsttarg.Directory.GetFiles())
            {
                _FirstDocumentItems.Add(new DocumentDisplay(d.FullName, d.Name ));
            }
        }
        private bool onPruefVmpbCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddPruefDoc) &&
                SelectedItem != null && _FirstDocumentItems.Any(x => x.DocumentType == 1)
                && _VmpbDocumentItems.All(x => x.DocumentType != 1);
        }

        private void onPruefVmpbExecuted(object obj)
        {
            var docu = VmpbInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag, SelectedItem.Vorgang]);
            VmpbInfo.Collect();
            var prefile = new FileInfo(_FirstDocumentItems.First(x => x.DocumentType == 1).FullName);
            var postfile = new FileInfo(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], SelectedItem.Auftrag + "_1.Gutteil"+ prefile.Extension));
            File.Copy(prefile.FullName, postfile.FullName, false);
            _VmpbDocumentItems.Add(new DocumentDisplay(postfile.FullName, postfile.Name));
        }
        private bool onVmpbCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddVmpb) &&
                SelectedItem != null && InWorkState == 0 && _VmpbDocumentItems.Any(x => x.DocumentType == 1);
        }
        private void onVmpbExecuted(object obj)
        {
            try
            {
                var mes = ProcessList.First(x => x.Auftrag == _SelectedValue) as VorgItem;
                var oa = new string[] { mes.Material.Trim(), mes.Auftrag, mes.Vorgang };
                var size = (string)obj;
      
                var docu = VmpbInfo.CreateDocumentInfos(oa);
                VmpbInfo.Collect();
                FileInfo vmFile;
                switch (size)
                {
                    case "size0":
                        vmFile = new FileInfo(docu[DocumentPart.Template]);
                        break;
                    case "size1":
                        vmFile = new FileInfo(docu[DocumentPart.Template_Size1]);
                        break;
                    case "size2":
                        vmFile = new FileInfo(docu[DocumentPart.Template_Size2]);
                        break;
                    case "size3":
                        vmFile = new FileInfo(docu[DocumentPart.Template_Size3]);
                        break;
                    case "size4":
                        vmFile = new FileInfo(docu[DocumentPart.Template_Size4]);
                        break;
                    default:
                        throw new NotImplementedException();
                }
                var vmtarg = new FileInfo(Path.Combine(docu[DocumentPart.OriginalFolder], docu[DocumentPart.File]));
                
    
                if (!vmtarg.Exists)
                {
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    File.Copy(vmFile.FullName, vmtarg.FullName.Trim());
                    var doku = db.Vorgangs.Single(x => x.Aid == mes.Auftrag && x.Vnr == short.Parse(mes.Vorgang));
                    doku.VorgangDocu ??= new VorgangDocu();
                    
                    doku.VorgangDocu.VmpbOriginal = vmtarg.Name.Trim();
                    doku.VorgangDocu.VmpbTemplate = vmFile.FullName.Trim();
                    InWorkState = doku.VorgangDocu.InWorkState = 1;
                    db.SaveChanges();
                    
                    var pi = new ProcessStartInfo(vmtarg.FullName.Trim())
                    {
                        UseShellExecute = true,
                        Verb = "OPEN"
                    };
                    Process.Start(pi);
                } 
            }
            catch (NotImplementedException)
            {
                _logger.LogWarning("{message}", "No Template definition");
                MessageBox.Show("Keine Vorlage definiert", "Vormusterprüfbericht", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("{message}", ex);
                MessageBox.Show(ex.Message, "Error Vmpb", MessageBoxButton.OK, MessageBoxImage.Error);
            }
 
        }
        private bool onVmpbOpenCanExecute(object arg)
        {
            return InWorkState > 0;
        }

        private void onVmpbOpenExecuted(object obj)
        {
            try
            {
                var oa = new string[] { SelectedItem.Material.Trim(), SelectedItem.Auftrag, SelectedItem.Vorgang };
                var docu = VmpbInfo.CreateDocumentInfos(oa);
                if (InWorkState > 0)
                {
                    var pi = new ProcessStartInfo(Path.Combine(docu[DocumentPart.OriginalFolder], docu[DocumentPart.File]))
                    {
                        UseShellExecute = true,
                        Verb = "OPEN"
                    };
                    Process.Start(pi);
                }
            }
            catch (FileNotFoundException e) 
            {
                MessageBox.Show(e.Message, "File IO", MessageBoxButton.OK, MessageBoxImage.Warning);
                _logger.LogWarning(e.ToString());
            }
            catch (Exception e)
            { 
                
                _logger.LogError(e.ToString());
            }
        }
        private bool onVmpbCreatePdfCanExecute(object arg)
        {
            return (SelectedItem.SourceVorgang?.VorgangDocu?.VmpbOriginal != null) && InWorkState == 1;
        }

        private void onVmpbCreatePdfExecuted(object obj)
        {

            try
            {
                if (SelectedItem.SourceVorgang?.VorgangDocu != null)
                {
                    var mes = ProcessList.First(x => x.Auftrag == _SelectedValue);
                    var oa = new string[] { mes.Material.Trim(), mes.Auftrag, SelectedItem.Vorgang };

                    var docu = VmpbInfo.CreateDocumentInfos(oa);
                    VmpbInfo.Collect();
                    var vmpFile = new FileInfo(SelectedItem.SourceVorgang?.VorgangDocu?.VmpbOriginal);
                    var path = Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath],
                        docu[DocumentPart.Folder], Path.GetFileNameWithoutExtension(vmpFile.Name));
                    Type officeType = Type.GetTypeFromProgID("Word.Application");

                    #if (officeType != null)

                        Microsoft.Office.Interop.Word.Application wordApp = new();
                        Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(SelectedItem.OrderDocu.VmpbOriginal);
                        wordDoc.ExportAsFixedFormat(path + ".pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        wordDoc.Close();
                        wordApp.Quit();
                    #endif

                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    db.VorgangDocus.Single(x => x.Vorgang.Aid == _SelectedValue).InWorkState = (int)(InWorkState = 2);
                    db.SaveChanges();
           
                    var docuItems = new DirectoryInfo(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], docu[DocumentPart.Folder]));
                    _PartDocumentItems.Clear();
                    foreach (var d in docuItems.GetFiles())
                    {
                        _PartDocumentItems.Add(new DocumentDisplay(d.FullName, d.Name));
                    }
                }
            }
            catch (FileNotFoundException e)
            {
                MessageBox.Show(e.Message, "Datei", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                _logger.LogWarning(e.ToString());
            }
            catch (COMException e)
            {
                MessageBox.Show("${e.Message}\nStellen Sie sicher dass die Datei nicht geöffnet ist.", "Datei", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                _logger.LogWarning(e.ToString());
                
            }
            catch (Exception e)
            {
                _logger.LogError(e.ToString());
            }
        }
        private bool onVmpbCloseCanExecute(object arg)
        {
            return InWorkState >= 2;
        }

        private void onVmpbCloseExecuted(object obj)
        {
            switch (InWorkState)
            {
                case 2: InWorkState = 3; break;
                case 3: InWorkState = 2; break;
            }           
        }
        private bool onVmpbDelCanExecute(object arg)
        {
            return SelectedItem.SourceVorgang?.VorgangDocu != null;
        }

        private void onVmpbDelExecuted(object obj)
        {
 
            if(SelectedItem.SourceVorgang?.VorgangDocu?.VmpbOriginal != null) File.Delete(SelectedItem.SourceVorgang.VorgangDocu.VmpbOriginal);
            
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var o = db.VorgangDocus.SingleOrDefault(x => x.Vorgang.Aid == SelectedValue);
            if (o != null)
            {
                db.VorgangDocus.Remove(o);
                db.SaveChanges();
            }
            InWorkState = 0;
            if (SelectedItem.SourceVorgang != null)
                SelectedItem.SourceVorgang.VorgangDocu = null;

        }
        private void OnOrderChanged()
        {
            _FirstDocumentItems.Clear();
            _VmpbDocumentItems.Clear();
            _PartDocumentItems.Clear();
            if (SelectedItem != null)
            {             
                var docu = FirstPartInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                string path = Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath]);
                if (Directory.Exists(path))
                {
                    foreach (var d in Directory.GetFiles(path).Where(x => x.Contains('~') == false))
                    {
                        FileInfo f = new FileInfo(d);
                        _FirstDocumentItems.Add(new DocumentDisplay(f.FullName, f.Name));
                    }
                }

                
                var vmdocu = VmpbInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                string vmpath = Path.Combine(docu[DocumentPart.RootPath], vmdocu[DocumentPart.SavePath]);
                if (Directory.Exists(vmpath))
                {
                    foreach (var d in Directory.GetFiles(vmpath).Where(x => x.Contains('~') == false))
                    {
                        FileInfo f = new FileInfo(d);
                        _VmpbDocumentItems.Add(new DocumentDisplay(f.FullName, f.Name));
                    }
                }
                var Mdocu = MeasureInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                string Mpath = Path.Combine(docu[DocumentPart.RootPath], Mdocu[DocumentPart.SavePath]);
                if (Directory.Exists(Mpath))
                {
                    foreach (var d in Directory.GetFiles(Mpath).Where(x => x.Contains('~') == false))
                    {
                        FileInfo f = new FileInfo(d);
                        _PartDocumentItems.Add(new DocumentDisplay(f.FullName, f.Name));                       
                    }
                }
            }
        }
        public void DragOver(IDropInfo dropInfo)
        {
            var t = (ItemsControl)dropInfo.VisualTarget;
            bool accept;
            switch (t.Name)
            {
                case "first":
                    accept = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddPruefDoc) && SelectedItem != null;
                    break;
                case "vmpb":
                    accept = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddVmpb) && SelectedItem != null;
                    break;
                case "part":
                    accept = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddMeasureDocu) && SelectedItem != null;
                    break;
                default: accept = false; break;
            }
            if (accept)
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.Effects = DragDropEffects.Copy;
            }
        }
        
        public void Drop(IDropInfo dropInfo)
        {
            if (dropInfo.Data is IDataObject f)
            {
                var o = (string[])f.GetData(DataFormats.FileDrop);
                
                if (o.Length > 0 && SelectedItem != null)
                {
                    var t = (ItemsControl)dropInfo.VisualTarget;
                    if (t.Name == "first")
                    {
                        var docu = FirstPartInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                        FirstPartInfo.Collect();
                        FileInfo source = new FileInfo(o[0]);
                        var target = new FileInfo(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], source.Name));
                        if (!source.Exists) { MessageBox.Show("Datei nicht vorhanden", "File IO Error", MessageBoxButton.OK, MessageBoxImage.Stop); return; }
                        if (!target.Exists)
                        {
                            File.Copy(source.FullName, target.FullName);
                            _FirstDocumentItems.Add(new DocumentDisplay(target.FullName, target.Name));
                        }
                    }
                    if (t.Name == "vmpb")
                    {
                        var docu = VmpbInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                        VmpbInfo.Collect();
                        FileInfo source = new FileInfo(o[0]);
                        var target = new FileInfo(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], source.Name));
                        if (!source.Exists) { MessageBox.Show("Datei nicht vorhanden", "File IO Error", MessageBoxButton.OK, MessageBoxImage.Stop); return; }
                        if (!target.Exists)
                        {
                            File.Copy(source.FullName, target.FullName);
                            _VmpbDocumentItems.Add(new DocumentDisplay(target.FullName, target.Name));
                        }
                    }
                    if (t.Name == "part")
                    {
                        var docu = MeasureInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
                        MeasureInfo.Collect();
                        FileInfo source = new FileInfo(o[0]);
                        var target = new FileInfo(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], source.Name));
                        if (!source.Exists) { MessageBox.Show("Datei nicht vorhanden", "File IO Error", MessageBoxButton.OK, MessageBoxImage.Stop); return; }
                        if (!target.Exists)
                        {
                            File.Copy(source.FullName, target.FullName);
                            _PartDocumentItems.Add(new DocumentDisplay() { FullName = target.FullName, Display = target.Name });
                        }
                    }

                }
            }
            //if (dropInfo.Data is DocumentDisplay d)
            //{
            //    var docu = VmpbInfo.CreateDocumentInfos([SelectedItem.Material, SelectedItem.Auftrag]);
            //    VmpbInfo.Collect();
            //    FileInfo source = new FileInfo(d.FullName);
            //    var target = new FileInfo(Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath], "vmpb" + source.Extension));
            //    if (!source.Exists) { MessageBox.Show("Datei nicht vorhanden", "File IO Error", MessageBoxButton.OK, MessageBoxImage.Stop); return; }
            //    if (!target.Exists)
            //    {
            //        File.Copy(source.FullName, target.FullName);
            //        _VmpbDocumentItems.Add(new DocumentDisplay() { FullName = target.FullName, Display = target.Name });
            //    }
            //}
        }

        public bool IsNavigationTarget(NavigationContext navigationContext)
        {
            return navigationContext.Parameters.GetValue<Vorgang>("order") != null;
        }

        public void OnNavigatedFrom(NavigationContext navigationContext)
        {
            //throw new NotImplementedException();
        }
        public void OnNavigatedTo(NavigationContext navigationContext)
        {
            var vrg = navigationContext.Parameters.GetValue<Vorgang>("order");
            if (vrg != null)
            {
                SelectedItem = new VorgItem(vrg);
                SelectedValue = _SelectedItem.Auftrag;
            }
        }
        public struct DocumentDisplay
        {
            
            public DocumentDisplay(string FullName, string Display) 
            { 
                _fullName = FullName;
                _display = Display;
                DocumentType = 0;
                if (Display.Contains("Gutteil"))
                {
                    DocumentType = 1;
                }
            }
            private string _fullName;
            public string FullName { get { return _fullName; } set { _fullName = value; } }
            private string _display;
            public string Display { get { return _display; } set { _display = value; } }
            public int DocumentType { get; }
        }
    }
}
