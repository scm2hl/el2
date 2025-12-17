using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using Prism.Dialogs;
using Prism.Events;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Threading;


namespace ModuleDeliverList.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class LieferViewModel : ViewModelBase, IViewModel
    {
        public ICollectionView OrdersView { get; private set; }
        public ICommand TextSearchCommand => _textSearchCommand ??= new RelayCommand(OnTextSearch);
        public ICommand FilterDeleteCommand => _filterDeleteCommand ??= new RelayCommand(OnFilterDelete);
        public ICommand ToggleFilterCommand => _toggleFilterCommand ??= new RelayCommand(OnToggleFilter);
        public ICommand SortAscCommand => _sortAscCommand ??= new RelayCommand(OnAscSortExecuted);
        public ICommand SortDescCommand => _sortDescCommand ??= new RelayCommand(OnDescSortExecuted);
        public ICommand ProjectPrioCommand { get; private set; }
        public ICommand CreateRtfCommand { get; private set; }
        public ICommand CreatePdfCommand { get; private set; }
        public ICommand CreateHtmlCommand { get; private set; }
        public ICommand SendMailCommand { get; private set; }
        public ICommand AttachmentCommand { get; private set; }
        private ConcurrentObservableCollection<Vorgang> _orders { get; } = [];
        private DB_COS_LIEFERLISTE_SQLContext DBctx { get; set; }
        public ActionCommand SaveCommand { get; private set; }
        public ActionCommand FilterSaveCommand { get; private set; }
        public ActionCommand InvisibilityCommand { get; private set; }
        public string Title { get; } = "Lieferliste";
        public bool HasChange => DBctx.ChangeTracker.HasChanges();
        private readonly ILogger _Logger;
        IDialogService _dialogService;
        private RelayCommand? _textSearchCommand;
        private RelayCommand? _filterDeleteCommand;
        private RelayCommand? _toggleFilterCommand;
        private RelayCommand? _sortAscCommand;
        private RelayCommand? _sortDescCommand;
        private string _searchFilterText = string.Empty;
        private string _selectedProjectFilter = string.Empty;
        private string _selectedSectionFilter = string.Empty;
        private string _markerCode = string.Empty;
        private readonly Lock _lock = new();
        private static System.Timers.Timer? _autoSaveTimer;
        private readonly IContainerProvider _container;
        private readonly IEventAggregator _ea;
        private readonly IUserSettingsService _settingsService;
        private CmbFilter _selectedDefaultFilter = CmbFilter.NOT_SET;
        private static readonly List<Ressource> _ressources = [];
        private static readonly SortedDictionary<int, string> _sections = [];
        public SortedDictionary<int, string> Sections => _sections;
        public List<string> PersonalFilterKeys { get; } = [.. PersonalFilterContainer.GetInstance().Keys];
        private ObservableCollection<ProjectStruct> _projects = [];
        public ObservableCollection<ProjectStruct> Projects
        {
            get { return _projects; }
            set
            {
                if (_projects != value)
                {
                    _projects = value;
                    NotifyPropertyChanged(() => Projects);
                }
            }
        }
        private bool _activeOnly = true;

        public bool ActiveOnly
        {
            get { return _activeOnly; }
            set
            {
                _activeOnly = value;
                OrdersView?.Refresh();
            }
        }

        private bool _filterInvers;
        public bool FilterInvers
        {
            get { return _filterInvers; }
            set
            {
                if (_filterInvers != value)
                {
                    _filterInvers = value;
                    NotifyPropertyChanged(() => FilterInvers);
                    OrdersView.Refresh();
                }
            }
        }
        public string SearchFilterText
        {
            get { return _searchFilterText; }
            set
            {
                if (_searchFilterText != value)
                {
                    _searchFilterText = value;
                    NotifyPropertyChanged(() => SearchFilterText);
                    OrdersView.Refresh();
                }
            }
        }
        public string MarkerCode
        {
            get { return _markerCode; }
            set
            {
                if (value != _markerCode)
                {
                    _markerCode = value;
                    NotifyPropertyChanged(() => MarkerCode);
                    OrdersView.Refresh();
                }
            }
        }
        private IApplicationCommands _applicationCommands;

        public IApplicationCommands ApplicationCommands
        {
            get => _applicationCommands;
            set
            {
                if (_applicationCommands != value)
                    _applicationCommands = value;
                NotifyPropertyChanged(() => ApplicationCommands);
            }
        }
        public NotifyTaskCompletion<ICollectionView>? OrderTask { get; private set; }

        internal CollectionViewSource OrdersViewSource { get; private set; } = new();

        public enum CmbFilter
        {
            [Description("Leer")]
            // ReSharper disable once InconsistentNaming
            NOT_SET = 0,
            [Description("ausgeblendet")]
            INVISIBLE,
            [Description("zum ablegen")]
            READY,
            [Description("Aufträge zum Starten")]
            START,
            [Description("Entwicklungsmuster (EM...)")]
            DEVELOP,
            [Description("Verkaufsmuster (VM...)")]
            SALES,
            [Description("Projekte mit Verzug")]
            PROJECTS_LOST,
            [Description("rote Aufträge")]
            ORDERS_RED,
            [Description("rote Projekte")]
            PROJECTS_RED,
            [Description("EXTERN")]
            EXERTN
        }
        public CmbFilter SelectedDefaultFilter
        {
            get => _selectedDefaultFilter;
            set
            {
                if (_selectedDefaultFilter != value)
                {
                    _selectedDefaultFilter = value;
                    NotifyPropertyChanged(() => SelectedDefaultFilter);
                    OrdersView.Refresh();
                }
            }
        }
        private string _selectedPersonalFilter = PersonalFilterContainer.GetInstance().Keys.First();

        public string? SelectedPersonalFilter
        {
            get
            {
                return _selectedPersonalFilter;
            }
            set
            {
                if (value != _selectedPersonalFilter)
                {
                    if (value != null)
                    {
                        _selectedPersonalFilter = value;
                        NotifyPropertyChanged(() => SelectedPersonalFilter);
                        OrdersView.Refresh(); 
                    }
                }
            }
        }
        private ProjectTypes.ProjectType _SelectedProjectType = ProjectTypes.ProjectType.None;

        public ProjectTypes.ProjectType SelectedProjectType
        {
            get { return _SelectedProjectType; }
            set
            {
                _SelectedProjectType = value;
                NotifyPropertyChanged(() => SelectedProjectType);
                OrdersView?.Refresh();
            }
        }

        public string SelectedProjectFilter
        {
            get => _selectedProjectFilter;
            set
            {
                if (value != _selectedProjectFilter)
                {
                    _selectedProjectFilter = value;
                    NotifyPropertyChanged(() => SelectedProjectFilter);
                    OrdersView?.Refresh();
                }
            }
        }

        public string SelectedSectionFilter
        {
            get { return _selectedSectionFilter; }
            set
            {
                if (_selectedSectionFilter != value)
                {
                    _selectedSectionFilter = value;
                    NotifyPropertyChanged(() => SelectedSectionFilter);
                    OrdersView?.Refresh();
                }
            }
        }
        public LieferViewModel(IContainerProvider container,
            IApplicationCommands applicationCommands,
            IEventAggregator ea,
            IUserSettingsService settingsService,
            IDialogService dialogService)
        {
            _applicationCommands = applicationCommands;
            DBctx = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            _container = container;
            var factory = _container.Resolve<ILoggerFactory>();
            _Logger = factory.CreateLogger<LieferViewModel>();

            try
            {
                _ea = ea;
                _settingsService = settingsService;
                _dialogService = dialogService;

                InvisibilityCommand = new ActionCommand(OnInvisibilityExecuted, OnInvisibilityCanExecute);
                SaveCommand = new ActionCommand(OnSaveExecuted, OnSaveCanExecute);
                FilterSaveCommand = new ActionCommand(OnFilterSaveExecuted, OnFilterSaveCanExecute);
                ProjectPrioCommand = new ActionCommand(OnSetProjectPrioExecuted, OnSetProjectPrioCanExecute);
                CreateRtfCommand = new ActionCommand(OnCreateRtfExecuted, OnCreateRtfCanExecute);
                CreateHtmlCommand = new ActionCommand(OnCreateHtmlExecuted, OnCreateHtmlCanExecute);
                AttachmentCommand = new ActionCommand(OnAttachmentExecuted, OnAttachmentCanExecute);
                OrderTask = new NotifyTaskCompletion<ICollectionView>(LoadDataAsync());

                _ea.GetEvent<MessageVorgangChanged>().Subscribe(MessageVorgangReceived);
                _ea.GetEvent<MessageOrderChanged>().Subscribe(MessageOrderReceived);
                _ea.GetEvent<MessageOrderEnclose>().Subscribe(MessageOrderEnclosed);
                _ea.GetEvent<EnableAutoSave>().Subscribe(AutoSaveEnable);
                SetAutoSave();
            }
            catch (Exception e)
            {
                _Logger.LogError(e.Message);
            }

        }

        private void AutoSaveEnable(bool obj)
        {
            if (obj)
            {
                if (_autoSaveTimer == null) SetAutoSave();
                _autoSaveTimer?.Start();
            }
            else
                _autoSaveTimer?.Stop();
        }

        private AbstracatBuilder CreateTableBuilder()
        {
            TableBuilder t = new();
            string[] headers = ["Auftragsnummer", "Material", "Bezeichnng", "Kurztext", "Termin"];
            AbstracatBuilder builder = new FlowTableBuilder(headers);
            List<Vorgang> query = OrdersView.Cast<Vorgang>().ToList();
            var sel = query.Select(x => new string?[]
            {
                x.Aid,
                x.AidNavigation.Material,
                x.AidNavigation.MaterialNavigation?.Bezeichng,
                x.Text,
                x.Termin.ToString(),
            }).ToList();
            builder.SetContext((List<string?[]>)sel);
            t.Build(builder);
            return builder;
        }

        // This method accepts an input stream and a corresponding data format.  The method
        // will attempt to load the input stream into a TextRange selection, apply Bold formatting
        // to the selection, save the reformatted selection to an alternat stream, and return 
        // the reformatted stream.  
        Stream BoldFormatStream(Stream inputStream, string dataFormat)
        {
            // A text container to read the stream into.
            FlowDocument workDoc = new FlowDocument();
            TextRange selection = new TextRange(workDoc.ContentStart, workDoc.ContentEnd);
            Stream outputStream = new MemoryStream();

            try
            {
                // Check for a valid data format, and then attempt to load the input stream
                // into the current selection.  Note that CanLoad ONLY checks whether dataFormat
                // is a currently supported data format for loading a TextRange.  It does not 
                // verify that the stream actually contains the specified format.  An exception 
                // may be raised when there is a mismatch between the specified data format and 
                // the data in the stream. 
                if (selection.CanLoad(dataFormat))
                    selection.Load(inputStream, dataFormat);
            }
            catch (Exception e) { return outputStream; /* Load failure; return a null stream. */ }

            // Apply Bold formatting to the selection, if it is not empty.
            if (!selection.IsEmpty)
                selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);

            // Save the formatted selection to a stream, and return the stream.
            if (selection.CanSave(dataFormat))
                selection.Save(outputStream, dataFormat);

            return outputStream;
        }

        private void MessageOrderEnclosed(OrderRb rb)
        {
            try
            {
                if (rb.Abgeschlossen)
                {
                    var o = _orders.Where(x => x.Aid == rb.Aid);
                    foreach (var x in o)
                    {
                        using (_lock.EnterScope())
                        {
                            _orders.Remove(x);
                            DBctx.ChangeTracker.Entries<OrderRb>().First(x => x.Entity.Aid == rb.Aid).State = EntityState.Unchanged;
                            OrdersView.Refresh();
                            _Logger.LogInformation("Auftrag abgeschlossen: {message}", rb.Aid);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "MsgReceivedEnclosed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void MessageOrderReceived(List<(string, string)?> rb)
        {
            try
            {
                Task.Run(() =>
                {
  
                    foreach ((string, string) rbId in rb.Where(x => x != null).Select(v => ((string, string))v))
                    {
                        _Logger.LogInformation("commin {message}, {1}", rbId.Item1, rbId.Item2);
                        if (_orders.Any(x => x.Aid == rbId.Item2))
                        {
                            using (_lock.EnterScope())
                            {
                                var ord = _orders.First(x => x.Aid == rbId.Item2).AidNavigation;
                                DBctx.Entry<OrderRb>(ord).Reload();
                                _Logger.LogInformation("reloaded {message}", ord.Aid);
                                ord.RunPropertyChanged();

                                foreach (var o in _orders.Where(x => x.Aid.Trim() == rbId.Item2))
                                {
                                    
                                    DBctx.Entry<Vorgang>(o).Reload();
                                    _Logger.LogInformation("reloaded {message}", o.VorgangId);

                                    o.RunPropertyChanged();
                                    DBctx.ChangeTracker.Entries<Vorgang>().First(x => x.Entity.VorgangId.Trim() == o.VorgangId.Trim()).State = EntityState.Unchanged;
                                }
                            }
                        }
                        else
                        {
                            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                            foreach (var v in db.Vorgangs.Where(x => x.Aid.Trim() == rbId.Item2.Trim()))
                            {
                                if (v.Aktuell)
                                    Application.Current.Dispatcher.Invoke(AddRelevantProcess, (rbId.Item1, v.VorgangId));                                 
                            }
                        }
                    }
                    
                });
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "MsgReceivedLieferlisteOrder", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void MessageVorgangReceived(List<(string, string)?> vrgIdList)
        {
            try
            {

                Task.Run(() =>
                {

                    foreach (var vrg in vrgIdList.Where(x => x != null))
                    {
                        Vorgang v;
                        _Logger.LogInformation("commin {message}, {1}", vrg.Value.Item1, vrg.Value.Item2);
                        if (vrg != null)
                        {
                            if (_orders.Any(x => x.VorgangId == vrg.Value.Item2))
                            {
                                using (_lock.EnterScope())
                                {
                                    v = _orders.Single(x => x.VorgangId == vrg.Value.Item2);
                                    DBctx.Entry<Vorgang>(v).Reload();
                                    _Logger.LogInformation("reloaded {message}", v.VorgangId);
                                    v.RunPropertyChanged();                                   
                                }
                            }
                            else
                            {
                                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                                var vv = db.Vorgangs.SingleOrDefault(x => x.VorgangId.Trim() == vrg.Value.Item2);

                                _Logger.LogInformation("maybe adding {message}", vv?.VorgangId);
                                var b = Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, AddRelevantProcess, vrg);
                            }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "MsgReceivedLieferlisteVorgang", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool OrdersView_FilterPredicate(object value)
        {

            try
            {
                var ord = (Vorgang)value;

                var accepted = (_activeOnly) ? ord.Aktuell : true;

                if (accepted && _selectedDefaultFilter == CmbFilter.NOT_SET) accepted = ord.Visability;

                if (!string.IsNullOrWhiteSpace(_searchFilterText))
                {
                    if (accepted)
                    {
                        if (!(accepted = ord.Aid.Contains(_searchFilterText, StringComparison.CurrentCultureIgnoreCase)))
                            if (!(accepted = ord.AidNavigation.Material?.Contains(_searchFilterText, StringComparison.CurrentCultureIgnoreCase) ?? false))
                                accepted = ord.AidNavigation.MaterialNavigation?.Bezeichng?.Contains(_searchFilterText, StringComparison.CurrentCultureIgnoreCase) ?? false;
                    }
                }

                if (accepted && _selectedDefaultFilter == CmbFilter.INVISIBLE) accepted = !ord.Visability == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.READY) accepted = ord.AidNavigation.Fertig == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.START)
                    accepted = (ord.Text?.Contains("STARTEN", StringComparison.CurrentCultureIgnoreCase) ?? false) == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.SALES) accepted = (ord.Aid.StartsWith("VM")) == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.DEVELOP) accepted = (ord.Aid.StartsWith("EM")) == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.EXERTN) accepted = (ord.ArbPlSap == "_EXTERN_") == !FilterInvers;
                if (accepted) accepted = !ord.AidNavigation.Abgeschlossen;
                if (accepted && _selectedProjectFilter != "_keine") accepted = ord.AidNavigation.ProId == _selectedProjectFilter;
                if (accepted && _SelectedProjectType != ProjectTypes.ProjectType.None) accepted = ord.AidNavigation.Pro?.ProjectType == (int)_SelectedProjectType;
                if (accepted && _selectedSectionFilter != "_keine") accepted = _ressources?
                        .FirstOrDefault(x => x.Inventarnummer == ord.ArbPlSap?[3..])?
                        .WorkArea?.Bereich == _selectedSectionFilter;
                if (accepted && _markerCode != string.Empty) accepted = ord.AidNavigation.MarkCode?.Contains(_markerCode, StringComparison.InvariantCultureIgnoreCase) ?? false;
                if (accepted && _selectedDefaultFilter == CmbFilter.PROJECTS_LOST) accepted = ProjectsLost(ord.AidNavigation.Pro) == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.ORDERS_RED) accepted = ord.AidNavigation.Prio?.Length > 0 == !FilterInvers;
                if (accepted && _selectedDefaultFilter == CmbFilter.PROJECTS_RED) accepted = ord.AidNavigation.Pro?.ProjectPrio == !FilterInvers;

                if (accepted && _selectedPersonalFilter != null)
                {
                    var b = PersonalFilterContainer.GetInstance();
                    accepted = (_selectedPersonalFilter != "_keine") ? b[_selectedPersonalFilter].TestValue(ord, _container) : true;
                }
                return accepted;
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.ToString(), "Filter Lieferliste", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private bool ProjectsLost(Project? pro)
        {
            if (pro != null)
            {
                foreach(var item in pro.OrderRbs.Where(x => x.Abgeschlossen == false))
                {
                    if(item.Vorgangs.Any(x => x.SpaetEnd < DateTime.Now && x.QuantityMiss > 0)) return true;
                }
            }
            return false;
        }

        private void OnTextSearch(object commandParameter)
        {
            if (commandParameter is string search)

                if (!search.IsNullOrEmpty() && search.Length >= 3)
                    SearchFilterText = search;
        }

        private void SetAutoSave()
        {
            _autoSaveTimer = new System.Timers.Timer(15000);
            _autoSaveTimer.Elapsed += OnAutoSave;
            _autoSaveTimer.AutoReset = true;
        }

        private void OnAutoSave(object? sender, ElapsedEventArgs e)
        {
            try
            {
                if (OrderTask != null && OrderTask.IsSuccessfullyCompleted)
                {

                    if (_lock.TryEnter())
                    {
                        if (DBctx.ChangeTracker.HasChanges()) DBctx.SaveChangesAsync();

                    }

                }
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "AutoSave", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (_lock.IsHeldByCurrentThread)
                    _lock.Exit();
            }
        }

        #region Commands

        private bool OnCreateRtfCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.CopyClipboard);
        }

        private void OnCreateRtfExecuted(object obj)
        {
            AbstracatBuilder Tbuilder = CreateTableBuilder();

            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            var sett = new UserSettingsService();
            dlg.InitialDirectory = string.IsNullOrEmpty(sett.PersonalFolder) ? Environment.GetFolderPath(Environment.SpecialFolder.Personal) :
                sett.PersonalFolder;
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".rtf"; // Default file extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dlg.FileName;
                using FileStream fs = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                FlowDocument flow = Tbuilder.GetDoc() as FlowDocument;
                TextRange tr = new TextRange(flow.ContentStart, flow.ContentEnd);
                tr.Save(fs, DataFormats.Rtf);
            }
        }
        private bool OnCreateHtmlCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.CopyClipboard);
        }

        private void OnCreateHtmlExecuted(object obj)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            var sett = new UserSettingsService();
            dlg.InitialDirectory = string.IsNullOrEmpty(sett.PersonalFolder) ? Environment.GetFolderPath(Environment.SpecialFolder.Personal) :
                sett.PersonalFolder;
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".html"; // Default file extension
            dlg.Filter = "Web documents (.html)|*.htm"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dlg.FileName;

                AbstracatBuilder Tbuilder = CreateTableBuilder();
                string content = Tbuilder.GetHtml();
                File.WriteAllText(filename, content);
            }
        }
        private bool OnAttachmentCanExecute(object arg)
        {
            return true;
        }

        private void OnAttachmentExecuted(object obj)
        {
            if (obj is Vorgang vrg)
            {
                var par = new DialogParameters();
                par.Add("vrg", vrg);
                _dialogService.ShowDialog("AttachmentDialog", par);
            }
        }
        private bool OnSetProjectPrioCanExecute(object arg)
        {
            if (arg is Vorgang v)
                if (v.AidNavigation.Pro != null)
                    return PermissionsProvider.GetInstance().GetUserPermission(Permissions.ProjectPrio);
            return false;
        }

        private void OnSetProjectPrioExecuted(object obj)
        {
            if (obj is Vorgang v)
            {
                if (v.AidNavigation.Pro != null)
                    v.AidNavigation.Pro.ProjectPrio = !v.AidNavigation.Pro.ProjectPrio;
            }
        }

        private bool OnInvisibilityCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.LieferVrgInvis);
        }

        private void OnInvisibilityExecuted(object obj)
        {
            if (obj is Vorgang v)
            {
                v.Visability = !v.Visability;
                OrdersView.Refresh();
            }
        }
        private bool OnFilterSaveCanExecute(object arg)
        {
            return false;
        }

        private void OnFilterSaveExecuted(object obj)
        {
            throw new NotImplementedException();
        }
        private void OnFilterDelete(object obj)
        {
            SearchFilterText = string.Empty;
            SelectedDefaultFilter = CmbFilter.NOT_SET;
            SelectedProjectType = ProjectTypes.ProjectType.None;
            SelectedProjectFilter = Projects.ElementAt(0).ProjectPsp;
            SelectedSectionFilter = Sections.ElementAt(0).Value;
            SelectedPersonalFilter = PersonalFilterContainer.GetInstance().Keys[0];
            MarkerCode = string.Empty;
            FilterInvers = false;
            _Logger.LogInformation("Filter Keys {0}", string.Join(",", PersonalFilterContainer.GetInstance().Keys));
        }
        private void OnToggleFilter(object obj)
        {
            FilterInvers = !FilterInvers;
        }

        private void OnSaveExecuted(object obj)
        {
            try
            {

                using (_lock.EnterScope());
                {
                    DBctx.SaveChanges();
                }
  
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
                MessageBox.Show(string.Format("{0}\nInnerEx\n{1}",e.Message,e.InnerException), "OnSave Liefer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private bool OnSaveCanExecute(object arg)
        {
            if (_settingsService.IsAutoSave) return false;
            try
            {
                bool res = false;
                if (OrderTask.IsSuccessfullyCompleted && _lock.TryEnter())
                {
                    res = DBctx.ChangeTracker.HasChanges();
                    
                }

                return res;
            }
            catch (InvalidOperationException e)
            {
                _Logger.LogError("{message}", e.ToString());
                return false;
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
                return false;
            }
            finally
            {
                if (_lock.IsHeldByCurrentThread)
                    _lock.Exit();
            }
        }

        private void OnDescSortExecuted(object parameter)
        {
            var v = Translate();
            if (v != string.Empty)
            {
                OrdersViewSource.SortDescriptions.Clear();
                OrdersViewSource.SortDescriptions.Add(new SortDescription(v, ListSortDirection.Descending));
                OrdersView.Refresh();
            }
        }

        private void OnAscSortExecuted(object parameter)
        {

            var v = Translate();

            if (v != string.Empty)
            {
                OrdersView.SortDescriptions.Clear();
                OrdersView.SortDescriptions.Add(new SortDescription(v, ListSortDirection.Ascending));
                OrdersView.Refresh();
            }
        }
        #endregion

        public string HasMouse { get; set; }
        private string Translate()
        {
            string v = HasMouse switch
            {
                "txtOrder" => "Aid",
                "txtVnr" => "Vorgangs.Vnr",
                "txtMatText" => "MaterialNavigation.Bezeichng",
                "txtMatTTNR" => "Material",
                "txtVrgText" => "Vorgangs.Text",
                "txtPlanT" => "Termin",
                "txtProj" => "ProId",
                "txtPojInfo" => "ProInfo",
                "txtQuant" => "Quantity",
                "txtWorkArea" => "Vorgangs.ArbplSap",
                "txtQuantYield" => "Vorgangs.QuantityYield",
                "txtScrap" => "Vorgangs.QuantityScrap",
                "txtOpen" => "Vorgangs.QuantityMiss",
                "txtRessTo" => "Vorgangs.RidNavigation.RessName",
                _ => string.Empty,
            };

            return v;
        }

        public async Task<ICollectionView> LoadDataAsync()
        {
            //_projects.Add(new Project() { ProjectPsp = "leer"});

            if (!_sections.ContainsKey(0)) _sections.Add(0, "_keine");
            var a = await DBctx.Vorgangs
               .Include(v => v.AidNavigation)
               .ThenInclude(x => x.MaterialNavigation)
               .Include(d => d.AidNavigation.DummyMatNavigation)
               .Include(r => r.RidNavigation)
               .Include(m => m.AidNavigation.DummyMatNavigation)
               .Include(p => p.AidNavigation.Pro)
               .Include(v => v.ArbPlSapNavigation)
               .Where(x => x.AidNavigation.Abgeschlossen == false && x.SysStatus != null && !x.SysStatus.Contains("RÜCK"))
               .ToListAsync();
            
            var attPro = await DBctx.ProjectAttachments.Select(x => x.ProjectPsp).ToListAsync();
            var attVrg = await DBctx.VorgangAttachments.Select(x => x.VorgangId).ToListAsync();
            var ress = await DBctx.Ressources.AsNoTracking()
                .Include(x => x.WorkArea)
                .ToArrayAsync();
            var filt = await DBctx.ProductionOrderFilters.AsNoTracking().ToArrayAsync();
            _ressources.AddRange(ress);
            SortedSet<ProjectStruct> pl = [new("_keine", ProjectTypes.ProjectType.None, string.Empty)];
            await Task.Factory.StartNew(() =>
            {
                HashSet<Vorgang> result = new();

                SortedDictionary<string, ProjectTypes.ProjectType> proj = new();
 
                    bool relev = false;
                foreach (var group in a.GroupBy(x => x.Aid))
                {
                    if (filt.Any(y => y.OrderNumber == group.Key))
                    {
                        relev = true;
                    }
                    else
                    {
                        foreach (var vorg in group)
                        {

                            if (vorg.ArbPlSap?.Length >= 3 && !relev)
                            {
                                if (int.TryParse(vorg.ArbPlSap[..3], out int c))
                                    if (UserInfo.User.AccountCostUnits.Any(y => y.CostId == c))
                                    {
                                        relev = true;
                                        break;
                                    }
                            }

                        }
                    }
                    if (relev)
                    {
                            
                        foreach (var vorg in group)
                        {
                            if (vorg.AidNavigation.ProId != null && !vorg.AidNavigation.ProId.StartsWith('0'))
                            {
                                var p = vorg.AidNavigation.Pro;
                                if (p != null)
                                {
                                    pl.Add(new(p.ProjectPsp.Trim(), (ProjectTypes.ProjectType)p.ProjectType, p.ProjectInfo));
                                    p.AttCount = attPro.Where(x => x == p.ProjectPsp).Count();
                                }
                            }


                            vorg.AttCount = attVrg.Where(x => x == vorg.VorgangId).Count();
                            result.Add(vorg);
                            var inv = (vorg.ArbPlSap != null) ? vorg.ArbPlSap[3..] : string.Empty;
                            var z = ress.FirstOrDefault(x => x.Inventarnummer?.Trim() == inv)?.WorkArea;
                            if (z != null)
                            {
                                if (!_sections.ContainsKey(z.Sort) && z.Bereich != null)
                                    _sections.Add(z.Sort, z.Bereich);
                            }
                            
                        }
                    }
                        
                    relev = false;                                              
                }
                _orders.AddRange(result.OrderBy(x => x.SpaetEnd));
                
            });
            Projects.AddRange(pl.OrderBy(x => x.ProjectType));
            SelectedProjectFilter = pl.ElementAt(0).ProjectPsp;
            SelectedSectionFilter = _sections.ElementAt(0).Value;
            OrdersView = CollectionViewSource.GetDefaultView(_orders);
            OrdersView.Filter += OrdersView_FilterPredicate;
            ICollectionViewLiveShaping? live = OrdersView as ICollectionViewLiveShaping;
            if (live != null)
            {
                if (live.CanChangeLiveFiltering)
                {
                    live.LiveFilteringProperties.Add("Aktuell");
                    live.LiveFilteringProperties.Add("AidNavigation.Abgeschlossen");
                    live.IsLiveFiltering = true;
                    live.IsLiveSorting = false;
                }
            }
            return OrdersView;
        }
        private bool AddRelevantProcess((string, string) income)
        {
            try
            {
                bool returnValue = false;
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var vrg = db.Vorgangs
                    .Include(x => x.AidNavigation)
                    .ThenInclude(x => x.MaterialNavigation)
                    .Include(x => x.AidNavigation.DummyMatNavigation)
                    .Include(x => x.RidNavigation)
                    .Single(x => x.VorgangId.Trim() == income.Item2.Trim());
                var vrgAdd = db.Vorgangs
                    .Where(x => x.Aid == vrg.Aid);
                _Logger.LogInformation("relevant? {message} {1}-{2} {3}", vrg.VorgangId, vrg.Aid, vrg.Vnr, vrg.ArbPlSap);
                foreach (var item in vrgAdd)
                {

                    if (item.ArbPlSap?.Length >= 3)
                    {
                        if (int.TryParse(item.ArbPlSap[..3], out int c))
                        {
                            if (UserInfo.User.AccountCostUnits.Any(y => y.CostId == c))
                            {
                                _orders.Add(vrg);
                                _Logger.LogInformation("added {message}", vrg.VorgangId);
                                
                                returnValue = true;
                                break;
                            }
                        }
                    }
                }
                
                return returnValue;
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                return false;
            }
        }
        public void Closing()
        {
            if (DBctx.ChangeTracker.HasChanges())
            {
                if (_settingsService.IsSaveMessage)
                {
                    var result = MessageBox.Show("Sollen die Änderungen in Lieferliste gespeichert werden?",
                        Title, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        DBctx.SaveChanges();
                    }
                }
                else DBctx.SaveChanges();
            }
        }
    }
    public class ProjectStruct(string ProjectPsp, ProjectTypes.ProjectType ProjectType, string? projectInfo) : IComparable
    {
        public string ProjectPsp { get; } = ProjectPsp;
        public ProjectTypes.ProjectType ProjectType { get; } = ProjectType;
        public string? ProjectInfo { get; } = projectInfo;

        public int CompareTo(object? obj)
        {
            if (obj == null) return 1;
            ProjectStruct? otherProjectStruct = obj as ProjectStruct;
            if (otherProjectStruct != null)
                return this.ProjectPsp.CompareTo(otherProjectStruct.ProjectPsp);
            else
                throw new ArgumentException("Object is not a ProjectStruct");
        }
    }
}
