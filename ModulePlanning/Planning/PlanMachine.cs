using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using GongSolutions.Wpf.DragDrop;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ModulePlanning.Specials;
using Prism.Dialogs;
using Prism.Events;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;


namespace ModulePlanning.Planning
{
    public interface IPlanMachineFactory
    {
        IContainerProvider Container { get; }
        IApplicationCommands ApplicationCommands { get; }
        IEventAggregator EventAggregator { get; }
        IUserSettingsService SettingsService { get; }
        IDialogService DialogService { get; }
    }
    public class PlanMachineFactory : IPlanMachineFactory
    {
        public IContainerProvider Container { get; }

        public IApplicationCommands ApplicationCommands { get; }

        public IEventAggregator EventAggregator { get; }

        public IUserSettingsService SettingsService { get; }

        public IDialogService DialogService { get; }
        public DB_COS_LIEFERLISTE_SQLContext DbContext { get; }

        public PlanMachineFactory(IContainerProvider container, IApplicationCommands applicationCommands,
            IEventAggregator eventAggregator, IUserSettingsService settingsService, IDialogService dialogService)
        {
            this.Container = container;
            this.ApplicationCommands = applicationCommands;
            this.EventAggregator = eventAggregator;
            this.SettingsService = settingsService;
            this.DialogService = dialogService;
            DataObject dt = new();
            
        }
        public PlanMachine CreatePlanMachine(Ressource res, DB_COS_LIEFERLISTE_SQLContext context)
        {
            return new PlanMachine(res, context, Container, ApplicationCommands, EventAggregator, SettingsService, DialogService);
        }

    }
    public interface IPlanMachine
    {
        public int Rid { get; }
    }
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public class PlanMachine : ViewModelBase, IPlanMachine, IDropTarget, IViewModel
    {

        #region Constructors

        public PlanMachine(Ressource ressource, DB_COS_LIEFERLISTE_SQLContext context, IContainerProvider container, IApplicationCommands applicationCommands,
            IEventAggregator eventAggregator, IUserSettingsService settingsService, IDialogService dialogService)
        {
            _container = container;
            var factory = container.Resolve<ILoggerFactory>();
            _logger = factory.CreateLogger<PlanMachine>();

            _rId = ressource.RessourceId;
            _applicationCommands = applicationCommands;
            _eventAggregator = eventAggregator;
            _settingsService = settingsService;
            _dialogService = dialogService;
            _db = context;
            Setup = settingsService.PlanedSetup;
            Initialize();
            LoadData(ressource);
            if(SelectedRadioButton != 0)
                CalculateEndTime();
            ProcessesCV.Refresh();

        }

        #endregion

        public bool Vis { get; private set; }

        public ICommand? SetMarkerCommand { get; private set; }
        public ICommand? OpenMachineCommand { get; private set; }
        public ICommand? MachinePrintCommand { get; private set; }
        public ICommand? HistoryCommand { get; private set; }
        public ICommand? FastCopyCommand { get; private set; }
        public ICommand? CorrectionCommand { get; private set; }
        public ICommand? NewCalculateCommand { get; private set; }
        public ICommand? NewStoppageCommand { get; private set; }
        public ICommand? DelStoppageCommand { get; private set; }
        private ILogger _logger;
        private readonly Lock _lock = new();
        private HashSet<string> ImChanged = [];
        private ShiftPlanService _shiftPlanService;
        private readonly int _rId;
        private string _title;
        public string Title => _title;
        public string Setup { get; }
        public bool HasChange => false; // _shifts.Any(x => x.IsChanged);
        public int Rid => _rId;
        private string? _name;
        public string? Name { get { return _name; }
            set { if (_name != value)
                {
                    _name = value;
                    NotifyPropertyChanged(() => Name);
                }
            }
        }
        private string? _description;
        
        public string? Description { get { return _description; }
            set { if (_description != value)
                {
                    _description = value;
                    NotifyPropertyChanged(() => Description);
                }
            }
        }
        private bool _isAdmin;
        public bool IsAdmin
        {
            get { return _isAdmin; }
            set
            {
                if (_isAdmin != value)
                {
                    _isAdmin = value;
                    NotifyPropertyChanged(() => IsAdmin);
                }
            }
        }
        private Dictionary<int, string> _ShiftCalendars = new() { { 0, "keine Berechnung" } };
        public Dictionary<int, string> ShiftCalendars => _ShiftCalendars;
        private Dictionary<int, string[]> _Stoppages = [];
        public Dictionary<int, string[]> Stoppages => _Stoppages;
        public ICollectionView StoppagesView { get; private set; }
        public ICollectionView ShiftCalendarView { get; private set; }
        private int _SelectedRadioButton;
        public int SelectedRadioButton
        {
            get { return _SelectedRadioButton; }
            set
            {
                if (value != _SelectedRadioButton)
                {
                    _SelectedRadioButton = value;
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    db.Ressources.First(x => x.RessourceId == _rId).ShiftCalendar = (_SelectedRadioButton == 0) ? null : _SelectedRadioButton;
                    db.SaveChanges();
                    _shiftPlanService ??= new ShiftPlanService(Rid, _container);
                    _shiftPlanService.ReloadShiftCalendar();
                }
            }
        }
        public string? InventNo { get; private set; }
        public WorkArea? WorkArea { get; set; }
        public List<int> CostUnits { get; set; } = [];
        public ObservableCollection<Vorgang>? Processes { get; set; }
        public ICollectionView ProcessesCV { get; private set; }
        public bool EnableRowDetails { get; private set; }
        private IContainerProvider _container;
        private IEventAggregator _eventAggregator;
        private IApplicationCommands? _applicationCommands;
        private IUserSettingsService _settingsService;
        private DB_COS_LIEFERLISTE_SQLContext _db;
        private readonly IDialogService _dialogService;
        private Vorgang _scrollItem;

        public Vorgang ScrollItem
        {
            get
            {
                return _scrollItem;
            }
            set
            {
                if (value != _scrollItem)
                {
                    _scrollItem = value;
                    NotifyPropertyChanged(() => ScrollItem);
                }
            }
        }

        public IApplicationCommands? ApplicationCommands
        {
            get { return _applicationCommands; }
            set
            {
                if (_applicationCommands != value)
                {
                    _applicationCommands = value;
                }
            }
        }
        internal CollectionViewSource ProcessesCVSource { get; set; } = new CollectionViewSource();


        private void LoadData(Ressource res)
        { 
            CostUnits.AddRange(res.RessourceCostUnits.Select(x => x.CostId));
            WorkArea = res.WorkArea;
            Name = res.RessName;
            _SelectedRadioButton = (res.ShiftCalendar == null) ? 0 : (int)res.ShiftCalendar;
            _title = res.Inventarnummer ?? string.Empty;
            Vis = res.Visability;
            Description = res.Info;
            InventNo = res.Inventarnummer;
            Processes = res.Vorgangs.ToObservableCollection();
            ProcessesCVSource.Source = Processes;
            ProcessesCVSource.SortDescriptions.Add(new SortDescription("SortPos", ListSortDirection.Ascending));
            ProcessesCVSource.IsLiveSortingRequested = true;
            //ProcessesCVSource.LiveSortingProperties.Add("SortPos");
            ProcessesCV = ProcessesCVSource.View;
            //ProcessesCV.SortDescriptions.Add(new SortDescription("SortPos", ListSortDirection.Ascending));
  
            ProcessesCV.CollectionChanged += OnProcessesChanged;
            var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            foreach(var s in db.ShiftCalendars.OrderBy(x => x.Id).AsNoTracking())
            {
                _ShiftCalendars.Add(s.Id, s.CalendarName);
            }
            foreach(var stop in db.Stopages.Where(x => x.Rid == Rid && x.Endtime > DateTime.Today).OrderBy(x => x.Starttime).AsNoTracking())
            {
                Stoppages.Add(stop.Id, [ stop.Description ?? "Stop Ausfall", string.Format("{0} - {1}",
                        stop.Starttime.ToString("d.M H:mm"), stop.Endtime.ToString("d.M H:mm"))]);
            }
            StoppagesView = CollectionViewSource.GetDefaultView(Stoppages);
            ShiftCalendarView = CollectionViewSource.GetDefaultView(ShiftCalendars);
        }

        private void OnProcessesChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {

        }

        private void Initialize()
        {
            SetMarkerCommand = new ActionCommand(OnSetMarkerExecuted, OnSetMarkerCanExecute);
            OpenMachineCommand = new ActionCommand(OnOpenMachineExecuted, OnOpenMachineCanExecute);
            HistoryCommand = new ActionCommand(OnHistoryExecuted, OnHistoryCanExecute);
            FastCopyCommand = new ActionCommand(OnFastCopyExecuted, OnFastCopyCanExecute);
            CorrectionCommand = new ActionCommand(OnCorrectionExecuted, OnCorrectionCanExecute);
            NewCalculateCommand = new ActionCommand(OnCalculateExecuted, OnCalculateCanExecute);
            NewStoppageCommand = new ActionCommand(OnNewStoppageExecuted, OnNewStoppageCanExecute);
            DelStoppageCommand = new ActionCommand(OnDelStoppageExecuted, OnDelStoppageCanExecute);
            _eventAggregator.GetEvent<MessageOrderChanged>().Subscribe(MessageOrderReceived);
            _eventAggregator.GetEvent<MessageVorgangChanged>().Subscribe(MessageVorgangReceived);
            _eventAggregator.GetEvent<SearchTextFilter>().Subscribe(MessageSearchFilterReceived);
            IsAdmin = PermissionsProvider.GetInstance().GetUserPermission(Permissions.AdminFunc);
            EnableRowDetails = _settingsService.IsRowDetails;           
        }

        private void CalculateEndTime()
        {
            try
            {
                if (_SelectedRadioButton == 0) return;
                _shiftPlanService ??= new Specials.ShiftPlanService(Rid, _container);
                DateTime start = DateTime.Now;
                foreach (var p in Processes.OrderBy(x => x.SortPos))
                {
                    var dur = GetProcessDuration(p);
                    var l = _shiftPlanService.GetEndDateTime(dur, start);
                    var diff = l.Subtract(start);

                    if (diff.TotalMinutes == 0) p.Extends = "---";
                    else
                    {
                        p.Extends = string.Format("{0}\n({1}){2:N2}min \n{3}",p.Responses.Max(x => x.Timestamp.ToString("d.MM.yy HH:mm"))
                            , p.QuantityMissNeo, dur, l.ToString("dd.MM.yy - HH:mm"));
                        p.Alert = (p.SpaetEnd != null) ? p.SpaetEnd.Value.Date < l.Date : false;
                        
                    }
                    p.RunPropertyChanged();
                    start = l;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
            }
        }
        private double GetProcessDuration(Vorgang vorgang)
        {
            try
            {
                double duration = 0.0;
                var r = vorgang.Rstze == null ? 0.0 : (short)vorgang.Rstze; //Setup time
                var c = vorgang.Correction == null ? 0.0 : (short)vorgang.Correction; //correction time
                if (vorgang.AidNavigation.Quantity != null && vorgang.AidNavigation.Quantity != 0.0) //if have Total quantity
                {
                    //calculation of the currently required time
                    var procT = vorgang.Beaze == null ? 0.0 : (short)vorgang.Beaze;
                    var quant = (short)vorgang.AidNavigation.Quantity;
                    var miss = vorgang.QuantityMissNeo == null ? 0.0 : (short)vorgang.QuantityMissNeo;
                    if (vorgang.Responses.Count == 0) 
                    {
                        duration = (procT + r + c) / quant * miss;
                        vorgang.Formula = string.Format("Formel: (Beaze + Rstze + Korrektur) / Menge * offen\n({0}+{1}+{2})/{3}*{4}",
                            procT,r,c,quant,miss);
                    }
                    else
                    {
                        duration = (procT + c) / quant * miss;
                        vorgang.Formula = string.Format("Formel: (Beaze + Korrektur) / Menge * offen\n({0}+{1})/{2}*{3}",
                            procT, c, quant, miss);
                    }
                }
                return duration;
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
                return 0.0;
            }
        }
        private void MessageSearchFilterReceived(string obj)
        {
            try
            {
                var ind = Processes?.LastOrDefault(x => x.Aid.Equals(obj));
                if (ind == null) ind = Processes?.LastOrDefault(x => x.AidNavigation?.Material?.Equals(obj) ?? false);
                if (ind == null) ind = Processes?.LastOrDefault(x => x.AidNavigation?.MaterialNavigation?.Bezeichng?.Equals(obj) ?? false);
                if (ind != null) ScrollItem = ind;
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
            }
        }
        private void MessageOrderReceived(List<(string, string)?> list)
        {
            try
            {
                Task.Run(() =>
                {

                    using (_lock.EnterScope())
                    {
                        if (Processes == null) { _lock.Exit(); return; }
                        if (list == null || list.Count == 0) { _lock.Exit(); return; }
                        foreach ((string, string)? item in list)
                        {
                            if (item == null) continue;
                            var pr = Processes.FirstOrDefault(x => x.VorgangId == item.Value.Item2);
                            if (pr != null)
                            {
    
                                if (_db.ChangeTracker.HasChanges()) { _db.SaveChanges(); }
                                
                                _db.Entry<Vorgang>(pr).Reload();
                                pr.RunPropertyChanged();
                                _logger.LogInformation("Planmachine - reloaded {message}", pr.VorgangId);
                            }
                            else
                            {
                                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                                var vo = db.Vorgangs
                                    .Include(x => x.AidNavigation)
                                    .ThenInclude(x => x.MaterialNavigation)
                                    .Include(x => x.AidNavigation.DummyMatNavigation)
                                    .Include(x => x.RidNavigation)
                                    .FirstOrDefault(x => x.VorgangId == item.Value.Item2);
                                if (vo?.SysStatus?.Contains("RÜCK") == false)
                                {
                                    Application.Current.Dispatcher.Invoke(new Action(() => Processes?.Add(vo)));
                                    _logger.LogInformation("PlanMachine - added {message} {1}-{2}", vo.VorgangId, vo.Aid, vo.Vnr);
                                }
                            }
                        }

                        foreach ((string, string)? item in list)
                        {
                            if (item == null) continue;
                            foreach (var v in Processes?.Where(x => x.Aid == item.Value.Item2))
                            {
                                using (_lock.EnterScope())
                                {
                                    if (_db.ChangeTracker.HasChanges()) { _db.SaveChanges(); }
                                }
                                _db.Entry<Vorgang>(v).Reload();
                                v.RunPropertyChanged();
                                _logger.LogInformation("Planmachine - reloaded {0} {1} {2}", v.VorgangId, v.Aid, v.Vnr);
                            }
                        }
                    }
                    
                });
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
            }
        }
        private void MessageVorgangReceived(List<(string, string)?> vorgangIdList)
        {
            try
            { 
                Task.Run(() =>
                {
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

                    foreach ((string, string)? idTuple in vorgangIdList)
                    {
                        if(idTuple != null)
                        {
                            var pr = Processes?.FirstOrDefault(x => x.VorgangId == idTuple.Value.Item2);
                            if (pr != null)
                            {

                                using (_lock.EnterScope())
                                {
                                    if (_db.ChangeTracker.HasChanges()) { _db.SaveChanges(); }
                                    _db.Entry<Vorgang>(pr).Reload();
                                    pr.RunPropertyChanged();

                                    _logger.LogInformation("PlanMachine - execute reload {message} {1}", pr.VorgangId, pr.SysStatus);
                                }

                                if ((pr.SysStatus?.Contains("RÜCK") ?? false) || pr.Rid == null)
                                {
                                    using (_lock.EnterScope())
                                    {
                                        pr.SortPos = "Z";
                                        Application.Current.Dispatcher.Invoke(new Action(() => Processes?.Remove(pr)));
                                        _logger.LogInformation("PlanMachine - removed {message} {1} - {2}", pr.VorgangId, pr.Aid, pr.Vnr);
                                        if (pr.Rid == null)
                                            _eventAggregator.GetEvent<MessagePlanmachineProcessRemoved>().Publish(pr);
                                        ProcessesCV.Refresh();
                                    }
                                }                                 
                            }
                            else if (db.Vorgangs.Find(idTuple.Value.Item2)?.Rid == Rid)
                            {
                                
                                var vo = db.Vorgangs
                                    .Include(x => x.AidNavigation)
                                    .ThenInclude(x => x.MaterialNavigation)
                                    .Include(x => x.AidNavigation.DummyMatNavigation)
                                    .Include(x => x.RidNavigation)
                                    .First(x => x.VorgangId == idTuple.Value.Item2);
                                _logger.LogInformation("PlanMachine - maybe add {message} {1}-{2}", vo.VorgangId, vo.Aid, vo.Vnr);
                                if (vo?.SysStatus?.Contains("RÜCK") == false)
                                {
                                    Application.Current.Dispatcher.Invoke(new Action(() => Processes?.Add(vo)));
                                    _logger.LogInformation("PlanMachine - added {message} {1}-{2}", vo.VorgangId, vo.Aid, vo.Vnr);
                                    ProcessesCV.Refresh();
                                }
                            }
                        }
                    }
                    
                    
                });        
            }
            catch (DbUpdateConcurrencyException ex)
            {
                //The code from Microsoft - Resolving concurrency conflicts 
                foreach (var entry in ex.Entries)
                {
 
                    var proposedValues = entry.CurrentValues; //Your proposed changes
                    var databaseValues = entry.GetDatabaseValues(); //Values in the Db

                    foreach (var property in proposedValues.Properties)
                    {
                        var proposedValue = proposedValues[property];
                        var databaseValue = databaseValues[property];
                        _logger.LogError("{message} {0} => DB {1}", ex.Message, proposedValue, databaseValue);
                        // TODO: decide which value should be written to database
                        // proposedValues[property] = <value to be saved>;
                    }

                    // Refresh original values to bypass next concurrency check
                    entry.OriginalValues.SetValues(databaseValues);

                }
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
                MessageBox.Show(string.Format("{0}\n{1}", ex.Message, ex.InnerException), "MsgReceivedPlanMachine", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ChangedValues(string invoker, Vorgang local, Vorgang remote)
        {
            if (invoker == "EL2")
            {
                bool changed = false;
                if (local.BemM != remote.BemM) { local.BemM = remote.BemM; changed = true; }
                if (local.BemMa != remote.BemMa) { local.BemMa = remote.BemMa; changed = true; }
                if (local.CommentMach != remote.CommentMach) { local.CommentMach = remote.CommentMach; changed = true; }
                if (local.BemT != remote.BemT) { local.BemT = remote.BemT; changed = true; }
                if (local.Bullet != remote.Bullet) { local.Bullet = remote.Bullet; changed = true; }
                if (local.Correction != remote.Correction) { local.Correction = remote.Correction; changed = true; }
                if (local.Rid != remote.Rid) { local.Rid = remote.Rid; changed = true; }
                if (local.SortPos != remote.SortPos) { local.SortPos = remote.SortPos; changed = true; }
                if (local.Termin != remote.Termin) { local.Termin = remote.Termin; changed = true; }

                if (changed) local.RunPropertyChanged();
            }
            else
            { 
                local.SpaetEnd = remote.SpaetEnd;
                local.SpaetStart = remote.SpaetStart; 
                local.QuantityScrap = remote.QuantityScrap;
                local.QuantityMiss = remote.QuantityMiss;
                local.QuantityMissNeo = remote.QuantityMissNeo;
                local.Aktuell = remote.Aktuell;
                local.QuantityYield = remote.QuantityYield;

                local.RunPropertyChanged();
            }
            
        }
        private bool OnCalculateCanExecute(object arg)
        {
            return _SelectedRadioButton != 0;
        }
        private void OnCalculateExecuted(object obj)
        {
            CalculateEndTime();
        }
        private bool OnDelStoppageCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.DelStoppage);
        }

        private void OnDelStoppageExecuted(object obj)
        {
            if(obj is int id)
            {
                Stoppages.Remove(id);
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var stop = db.Stopages.SingleOrDefault(x => x.Id == id);
                if (stop != null)
                {
                    db.Stopages.Remove(stop);
                    db.SaveChanges();

                    StoppagesView.Refresh();
                    _shiftPlanService.ReloadStoppage();
                }
            }
        }

        private bool OnNewStoppageCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddStoppage);
        }

        private void OnNewStoppageExecuted(object obj)
        {

            _dialogService.Show("InputStoppage", OnStoppageCallback);
        }

        private void OnStoppageCallback(IDialogResult result)
        {
            if(result.Result == ButtonResult.OK)
            {
                var stop = result.Parameters.GetValue<Stopage>("Stopage");
                if (stop != null)
                {
                    stop.Rid = Rid;
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    db.Stopages.Add(stop);
                    db.SaveChanges();

                    Stoppages.Add(stop.Id, [ stop.Description ?? "Stop Ausfall", string.Format("{0} - {1}",
                        stop.Starttime.ToString("d.M H:mm"), stop.Endtime.ToString("d.M H:mm"))]);
                    StoppagesView.Refresh();
                    _shiftPlanService.ReloadStoppage();
                }
            }
        }

 
        private bool OnCorrectionCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.Correction);
        }

        private void OnCorrectionExecuted(object obj)
        {
            if(obj is Vorgang vrg)
            {
                var par = new DialogParameters();
                par.Add("correction", vrg);
                _dialogService.ShowDialog("CorrectionDialog", par, CorrectionCallback);
                
            }
        }

        private void CorrectionCallback(IDialogResult result)
        {
            if (result.Result == ButtonResult.OK)
            {
                var vrg = result.Parameters.GetValue<Vorgang>("correction");
                var corr = result.Parameters.GetValue<double?>("correct");
                var v = Processes.First(x => x.VorgangId == vrg.VorgangId);
                v.Correction = (float?)corr;
                v.RunPropertyChanged();
                _logger.LogInformation("{message} {1}", "Set Correction:", v.Correction);
            }          
        }

        private bool OnHistoryCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.HistoryDialog);
        }

        private void OnHistoryExecuted(object obj)
        {
            try
            {
                if (obj is Vorgang vrg)
                {
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    var matInfo = db.OrderRbs.AsNoTracking()
                        .Include(x => x.MaterialNavigation)
                        .Include(x => x.DummyMatNavigation)
                        .Include(x => x.Vorgangs)
                        .Where(x => x.Material == vrg.AidNavigation.Material && x.Aid != vrg.Aid)
                        .ToList();

                    if (matInfo != null && matInfo.Count > 0)
                    {
                        var par = new DialogParameters();
                        par.Add("orderList", matInfo);
                        par.Add("VNR", vrg.Vnr);
                        par.Add("VID", vrg.VorgangId);
                        _dialogService.Show("HistoryDialog", par, HistoryCallBack);
                    }
                    else
                    {
                        MessageBox.Show("Keine Einträge vorhanden!", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    _logger.LogInformation("{message} {1}", "Call History:", vrg.VorgangId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
            }
        }

        private void HistoryCallBack(IDialogResult result)
        {
            try
            {
                if (result.Result == ButtonResult.Yes)
                {
                    string[] bemt = [];
                    var vid = result.Parameters.GetValue<string>("VID");
                    var bem = result.Parameters.GetValue<string>("Comment");
                    if (bem != null) bemt = bem.Split((char)29);
                    if (bemt.Length > 1)
                    {
                        var pr = Processes?.First(x => x.VorgangId == vid);
                        {
                            pr.BemT = String.Format("[{0}-{1}]{2}{3}",
                            UserInfo.User.UserId, DateTime.Now.ToShortDateString(), (char)29, bemt[1]);
                            pr.RunPropertyChanged();
                        }
                    }
                    _logger.LogInformation("{message} {1}", "History Callback:", bemt.ToString());
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
            }
        }
        private bool OnFastCopyCanExecute(object arg)
        {
            return arg is Vorgang && PermissionsProvider.GetInstance().GetUserPermission(Permissions.FastCopy);
        }

        private void OnFastCopyExecuted(object obj)
        {
            try
            {
                using (_lock.EnterScope())
                {
                    if (obj is Vorgang v)
                    {
                        var m = v.AidNavigation.Quantity.ToString();
                        var a = v.Aid.Trim();
                        var vnr = v.Vnr.ToString();
                        var mat = v.AidNavigation.Material?.Trim();
                        var bez = v.AidNavigation.MaterialNavigation?.Bezeichng?.Trim();

                        OnFastCopyExecuted(m);
                        OnFastCopyExecuted(bez ?? a);
                        OnFastCopyExecuted(mat ?? "DUMMY");
                        OnFastCopyExecuted(vnr);
                        OnFastCopyExecuted(a);
                    }
                    if (obj is string s)
                    { setTextToClipboard(s); }
                }
            }
            catch (Exception e)
            {
                _logger.LogError("{message}", e.ToString());
                MessageBox.Show(string.Format("{0}\n{1}", e.Message, e.InnerException), "FastCopy", MessageBoxButton.OK, MessageBoxImage.Error);
            }
  
        }

        private void setTextToClipboard(string text)
        {
            System.Windows.Clipboard.SetText(text);
            Task.Delay(500).Wait();          
        }

        private static bool OnSetMarkerCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.SETMARK);
        }

        private void OnSetMarkerExecuted(object obj)
        {
            try
            {
                if (obj == null) return;
                var values = (object[])obj;
                var name = (string)values[0];
                var desc = (Vorgang)values[1];

                if (desc != null)
                {
                    if (name == "DelBullet") desc.Bullet = Brushes.White.ToString();
                    if (name == "Bullet1") desc.Bullet = Brushes.Red.ToString();
                    if (name == "Bullet2") desc.Bullet = Brushes.Green.ToString();
                    if (name == "Bullet3") desc.Bullet = Brushes.Yellow.ToString();
                    if (name == "Bullet4") desc.Bullet = Brushes.Blue.ToString();

                    desc.RunPropertyChanged();
                    _logger.LogInformation("{message} {1}", "Set Bullet:", desc.Bullet);
                }
            }
            catch (System.Exception e)
            {
                _logger.LogError("{message}", e.ToString());
                MessageBox.Show(e.Message, "SetMarker", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private bool OnOpenMachineCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenMach);
        }

        private void OnOpenMachineExecuted(object obj)
        {
            try
            {
                var par = new DialogParameters();
                par.Add("PlanMachine", this);
                ListCollectionView lv = ProcessesCV as ListCollectionView;
                if(lv.IsEditingItem) lv.CommitEdit();
                if(lv.IsAddingNew) lv.CommitNew();
                _dialogService.Show("MachineView", par, MachineViewCallBack);
            }
            catch (Exception e)
            {
                _logger.LogError("{message}", e.ToString());
                MessageBox.Show(e.Message, "Error OpenMachine", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MachineViewCallBack(IDialogResult result)
        {
            if (result.Result == ButtonResult.Yes)
            {
                string[] bemt = [];
                var vid = result.Parameters.GetValue<string>("VID");
                var bem = result.Parameters.GetValue<string>("Comment");
            }
        }

 
        private void InsertItems(Vorgang Item, ListCollectionView Source, ListCollectionView Target, int Index, bool sorting)
        {

            try
            {
                Item.Rid = _rId;
                List<Vorgang> lst = [];
                //Target.IsLiveSorting = false;
                var p = Target.SourceCollection as Collection<Vorgang>;
                lst.AddRange(p.OrderBy(x => x.SortPos));
                int oldIndex = Target.IndexOf(Item);
                if (oldIndex == -1) // commes from outside
                {
                    if (Index >= Target.Count)
                    {
                        ((IList)Target.SourceCollection).Add(Item);
                        lst.Add(Item);
                    }
                    else
                    {
                        ((IList)Target.SourceCollection).Insert(Index, Item);
                        lst.Insert(Index, Item);
                    }
                    Source.Remove(Item);
                }

                if (oldIndex != -1) //sorting inside
                {
                    lst.RemoveAt(oldIndex);

                    if (Index > oldIndex) Index--;
                    // the actual index could have shifted due to the removal

                    lst.Insert(Index, Item);
                }
                var larr = lst.ToArray();
                
                for (int i = 0; i < larr.Length; i++)
                {
                    var vrg = p.First(x => x.Equals(larr[i]));
                    var oldSort = vrg.SortPos;
                    
                    vrg.SortPos = string.Format("{0,4:0}_{1,3:0}", Rid.ToString("D3"), i.ToString("D3"));
                    _logger.LogInformation("new sort {message} {0} ==> {1}",vrg.VorgangId, oldSort, vrg.SortPos);
                }
                Target.MoveCurrentTo(Item);
                
                if (WorkArea != null && WorkArea.CreateFolder)
                {
                    var work = _container.Resolve<WorkareaDocumentInfo>();
                    if (!string.IsNullOrEmpty(Item.AidNavigation.Material))
                    {
                        string?[] oa = [Item.AidNavigation.Material, Item.Aid, WorkArea.Bereich];

                        _ = work.CreateDocumentInfos(oa);
                        work.Collect();
                    }
                    else if(Item.AidNavigation.DummyMat != null)
                    {
                        string?[] oa = [Item.AidNavigation.DummyMat, Item.Aid, WorkArea.Bereich];

                        _ = work.CreateDocumentInfos(oa, true);
                        work.Collect();
                    }
                }
                Target.IsLiveSorting = true;
            }
            catch (Exception ex)
            {
                _logger.LogError("{message}", ex.ToString());
            }
        }
        public void Drop(IDropInfo dropInfo)
        {

            try
            {

                var s = dropInfo.DragInfo.SourceCollection as ListCollectionView;
                var t = dropInfo.TargetCollection as ListCollectionView;

                var v = dropInfo.InsertIndex;
                if (s != null && t != null)
                {
                    ListCollectionView? lv = ProcessesCV as ListCollectionView;
                    if (lv == null) return;
                    if (lv.IsAddingNew) { lv.CommitNew(); }
                    if (lv.IsEditingItem) { lv.CommitEdit(); }
                    if (dropInfo.Data is List<dynamic> vrgList)
                    {
                        foreach (var vrg in vrgList)
                        {
                            InsertItems(vrg, s, t, v, false);

                        }
                        _logger.LogInformation("{message} {id}", "Drops", vrgList.ToString());
                    }
                    else if (dropInfo.Data is Vorgang vrg)
                    {
                        InsertItems(vrg, s, t, v, dropInfo.IsSameDragDropContextAsSource);
                        _logger.LogInformation("{message} {id} {sort}", "Drops", vrg.VorgangId, vrg.SortPos);
    
                    }
                        
                    ProcessesCV.Refresh();
                }
                              
            }
            catch (DbUpdateConcurrencyException ex)
            {
                //The code from Microsoft - Resolving concurrency conflicts 
                foreach (var entry in ex.Entries)
                {
                    var proposedValues = entry.CurrentValues; //Your proposed changes
                    var databaseValues = entry.GetDatabaseValues(); //Values in the Db
                    if (databaseValues != null)
                    {
                        foreach (var property in proposedValues.Properties)
                        {
                            var proposedValue = proposedValues[property];
                            var databaseValue = databaseValues[property];
                            _logger.LogError("{message} {0} => DB {1}", ex.Message, proposedValue, databaseValue);
                            // TODO: decide which value should be written to database
                            // proposedValues[property] = <value to be saved>;
                        }

                        // Refresh original values to bypass next concurrency check
                        entry.OriginalValues.SetValues(databaseValues);
                    }
                    else { _logger.LogError("databaseValues == null Planmachine Drop DbUpdateConcurrencyException"); }
                    }
            }
            catch (ArgumentNullException e)
            {
                _logger.LogError("{message} \nSource: {source} \nData:\n{Parameter}", e.ToString(), e.Source, e.Data);
            }
            catch (Exception e)
            {
                _logger.LogError("{message}", e.ToString());
                string str = string.Format(e.Message + "\n" + e.InnerException);
                MessageBox.Show(str, "ERROR", MessageBoxButton.OK);
            }
        }

        public void DragOver(IDropInfo dropInfo)
        {
            if (PermissionsProvider.GetInstance().GetRelativeUserPermission(Permissions.MachDrop, WorkArea?.WorkAreaId ?? 0))
            {
                if (dropInfo.Data is Vorgang || dropInfo.Data is List<dynamic>)
                {
                    dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                    dropInfo.Effects = DragDropEffects.Move;
                }               
            }
        }

        void IViewModel.Closing()
        {
 
        }

        internal void SaveAll()
        {
            //using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            //foreach (var s in _shifts.Where(x => x.IsChanged))
            //{
                
            //    RessourceWorkshift ws = new RessourceWorkshift() { Rid = Rid, Sid = s.Shift.Sid };
            //    if (s.IsCheck)
            //        db.RessourceWorkshifts.Add(ws);
            //    else db.RessourceWorkshifts.Remove(ws);
            //    s.IsChanged = false;
            //}
            //db.SaveChanges();
            
        }
    }
}
