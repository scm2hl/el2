
using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using GongSolutions.Wpf.DragDrop;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ModulePlanning.Planning;
using Prism.Events;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;


namespace ModulePlanning.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class MachinePlanViewModel : ViewModelBase, IDropTarget, IViewModel
    {
        public string Title { get; } = "Teamleiter Zuteilung";
        public bool HasChange => Changed();
        private string searchArbpl = string.Empty;
        public string SearchArbpl
        {
            get { return searchArbpl; }
            set
            {
                if (searchArbpl != value)
                {
                    searchArbpl = value;
                    NotifyPropertyChanged(() => SearchArbpl);
                    ProcessCV.Refresh();
                }
            }
        }
        private RelayCommand? _selectionChangeCommand;
        private RelayCommand? _textSearchCommand;
        public ICommand SelectionChangeCommand => _selectionChangeCommand ??= new RelayCommand(SelectionChange);
        public ICommand TextSearchCommand => _textSearchCommand ??= new RelayCommand(OnTextSearch);
        public ICommand SaveCommand { get; private set; }
        public List<WorkArea> WorkAreas { get; } = [];
        private ConcurrentObservableCollection<PlanMachine> _machines { get; } = new();
        private ObservableCollection<Vorgang>? Priv_processes { get; set; }
        private ObservableCollection<Vorgang>? Priv_parking { get; set; }
        private DB_COS_LIEFERLISTE_SQLContext _DbCtx;
        private IContainerExtension _container;
        private IEventAggregator _ea;
        private IUserSettingsService _settingsService;
        public ICollectionView RessCV { get; private set; }
        public ICollectionView ProcessCV { get { return ProcessViewSource.View; } }
        public ICollectionView ParkingCV { get { return ParkingViewSource.View; } }
        private IApplicationCommands _applicationCommands;
        private static System.Timers.Timer? _autoSaveTimer;
        private NotifyTaskCompletion<ICollectionView>? _machineTask;

        public NotifyTaskCompletion<ICollectionView>? MachineTask
        {
            get { return _machineTask; }
            set
            {
                if (_machineTask != value)
                {
                    _machineTask = value;
                    NotifyPropertyChanged(() => MachineTask);
                }
            }
        }

        public IApplicationCommands ApplicationCommands
        {
            get { return _applicationCommands; }
            set
            {
                if (_applicationCommands != null)
                {
                    _applicationCommands = value;
                    NotifyPropertyChanged(() => ApplicationCommands);
                }
            }
        }
        public double SizePercent { get; }
        private readonly ILogger _Logger;
        private int _currentWorkArea;
        private string? _searchFilterText;
        private readonly Lock _lock = new();

        internal CollectionViewSource ProcessViewSource { get; } = new();
        internal CollectionViewSource ParkingViewSource { get; } = new();

        public MachinePlanViewModel(IContainerExtension container, IApplicationCommands applicationCommands, IEventAggregator ea, IUserSettingsService settingsService)
        {
            _container = container;
            var loggerFactory = _container.Resolve<ILoggerFactory>();
            _Logger = loggerFactory.CreateLogger<MachinePlanViewModel>();

            _applicationCommands = applicationCommands;
            _ea = ea;
            _settingsService = settingsService;
            _DbCtx = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            SizePercent = 340 * _settingsService.SizePercent / 100;
            SaveCommand = new ActionCommand(OnSaveExecuted, OnSaveCanExecute);

            LoadWorkAreas();
            SetAutoSaveTimer();
            MachineTask = new NotifyTaskCompletion<ICollectionView>(LoadMachinesAsync());

            SubscribeEventsAsync();

        }

        private void MessagePlug(Vorgang vorgang)
        {
            using (_lock.EnterScope())
            {
                Application.Current.Dispatcher.InvokeAsync(() => Priv_processes?.Add(vorgang));

                _Logger.LogInformation("user has removed from machine {message}-{0}", vorgang.Aid, vorgang.Vnr);
            }          
        }

        private void AutoSaveEnable(bool obj)
        {
            if (obj)
            {
                if (_autoSaveTimer == null)
                {
                    SetAutoSaveTimer();
                }
                _autoSaveTimer?.Start();
            }
            else
                _autoSaveTimer?.Stop();
        }

        private bool IsRelevant(string? ArbPl)
        {

            if (int.TryParse(ArbPl?[..3], out var cid))
            {
                return UserInfo.User.AccountCostUnits.Any(x => x.CostId == cid);
            }
            return false;
        }
        private void MessageVorgangReceived(List<(string, string)?> list)
        {
            _ = Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    using (_lock.EnterScope())
                    {
                        foreach (var item in list)
                        {
                            if (!item.HasValue)
                                continue;

                            var vo = Priv_processes?.FirstOrDefault(x => x.VorgangId == item.Value.Item2);
                            if (vo == null)
                                continue;

                            var entry = _DbCtx.Entry(vo);
                            if (entry.State == EntityState.Modified)
                                _DbCtx.SaveChanges();
                            entry.Reload();
                            vo.RunPropertyChanged();
                            _Logger.LogInformation("reloaded {message}-{0} rid {1} {2}", vo.Aid, vo.Vnr, vo.Rid, item.Value.Item1);

                            if (item.Value.Item1 == "EL2" && vo.Rid != null)
                            {
                                Priv_processes?.Remove(vo);
                                _Logger.LogInformation("user has added to machine {message}-{0} rid {1}", vo.Aid, vo.Vnr, vo.Rid);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _Logger.LogError("{message}", ex.ToString());
                }
            });
        }

        private void MessageOrderReceived(List<(string, string)?> list)
        {
            if (list == null || list.Count == 0)
                return;

            _ = Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                try
                {

                    foreach (var item in list)
                    {
                        if (!item.HasValue)
                            continue;

                        var aid = item.Value.Item2;
                        var proc = await GetVorgangsAsync(aid);
                        var existing = Priv_processes?
                            .Where(x => x.Aid == aid)
                            .ToArray() ?? Array.Empty<Vorgang>();

                        foreach (var process in proc.Concat(existing).DistinctBy(x => x.VorgangId))
                        {
                            var entry = _DbCtx.Entry(process);
                            if (entry.State == EntityState.Modified)
                                _DbCtx.SaveChanges();

                            entry.Reload();
                            if (entry.State == EntityState.Detached)
                            {
                                Priv_processes?.Remove(process);
                                continue;
                            }

                            process.RunPropertyChanged();
                            if (!IsRelevant(process.ArbPlSap) || process.Rid != null ||
                                (process.SysStatus?.Contains("RÜCK") ?? false))
                            {
                                Priv_processes?.Remove(process);
                                continue;
                            }

                            if (Priv_processes?.Any(x => x.VorgangId == process.VorgangId) != true)
                            {
                                Priv_processes?.Add(process);
                                _Logger.LogInformation("scope pool added {message}-{0}", process.Aid, process.Vnr);
                            }
                        }

                    }
                    
                }
                catch (Exception ex)
                {
                    _Logger.LogError("{message}", ex.ToString());
                    MessageBox.Show(ex.Message, "MessageOrderReceived MachinePlan", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }
 
        private bool OnSaveCanExecute(object arg)
        {
            
            if (_settingsService.IsAutoSave) return false;
            try
            {
                bool res = false;
                if (MachineTask.IsSuccessfullyCompleted && _lock.TryEnter())
                {
                    res = _DbCtx.ChangeTracker.HasChanges();
                    _lock.Exit();
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
        }

        private void OnSaveExecuted(object obj)
        {
            try
            {
                using (_lock.EnterScope())
                {
                    _DbCtx.SaveChanges();
                }
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
                        _Logger.LogError("{message} {0} => DB {1}", ex.Message, proposedValue, databaseValue);
                        // TODO: decide which value should be written to database
                        // proposedValues[property] = <value to be saved>;
                    }

                    // Refresh original values to bypass next concurrency check
                    entry.OriginalValues.SetValues(databaseValues);

                }
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
                //MessageBox.Show(e.ToString(), "OnSave MachPlan", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void LoadWorkAreas()
        {
            foreach (var accWork in UserInfo.User.AccountWorkAreas.OrderByDescending(x => x.Standard))
            {
                WorkAreas.Add(accWork.WorkArea);
            }           
        }
        private async Task<List<Vorgang>> GetVorgangsAsync(string? aid, IReadOnlyCollection<int>? resourceIds = null)
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var query = db.Vorgangs
                .Include(x => x.AidNavigation)
                .ThenInclude(x => x.OrderComponents)
                .ThenInclude(x => x.MaterialNavigation)
                .Include(x => x.AidNavigation.DummyMatNavigation)
                .Include(x => x.AidNavigation.MaterialNavigation)
                .Include(x => x.ArbPlSapNavigation)
                .Include(x => x.Responses)
                .Include(x => x.RidNavigation)
                .ThenInclude(x => x!.WorkArea)
                .Include(x => x.RidNavigation)
                .ThenInclude(x => x!.RessourceCostUnits)
                .Where(y => y.SysStatus != null
                    && y.Text != null
                    && y.ArbPlSapNavigation!.Ressource!.WorkAreaId != 5
                    && y.Text.ToLower().Contains("auftrag starten") == false
                    && y.SysStatus.Contains("RÜCK") == false);

            query = aid != null
                ? query.Where(y => y.Aid == aid)
                : query.Where(y => y.AidNavigation.Abgeschlossen == false);

            if (resourceIds != null)
            {
                query = query.Where(y =>
                    (y.Rid.HasValue && resourceIds.Contains(y.Rid.Value)) ||
                    (y.ArbPlSapNavigation!.RessourceId.HasValue && resourceIds.Contains(y.ArbPlSapNavigation.RessourceId.Value)));
            }

            return await query.AsSplitQuery().ToListAsync();

        }

        private void SetAutoSaveTimer()
        {
            _autoSaveTimer = new System.Timers.Timer(15000);
            _autoSaveTimer.Elapsed += OnAutoSaveAsync;
            _autoSaveTimer.AutoReset = true;
            _autoSaveTimer.Enabled = true;
        }

        private void OnAutoSaveAsync(object? sender, ElapsedEventArgs e)
        {
            if (!Application.Current.Dispatcher.CheckAccess())
            {
                _ = Application.Current.Dispatcher.InvokeAsync(() => OnAutoSaveAsync(sender, e));
                return;
            }

            try
            {
                if (MachineTask != null && MachineTask.IsSuccessfullyCompleted && _lock.TryEnter())
                {
                    try
                    {
                        
                        if (_DbCtx.ChangeTracker.Entries().All(x => x.State == EntityState.Detached))
                        {
                            _Logger.LogInformation("AutoSave MachPlan: {message}", "All Detached.");
                        }
                        if (_DbCtx.ChangeTracker.HasChanges())
                            _DbCtx.SaveChanges();
                    }
                    finally
                    {
                        _lock.Exit();
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "AutoSave MachPlan", MessageBoxButton.OK, MessageBoxImage.Error);
                _Logger.LogError("{message}", ex.ToString());
            }
        }
        private async Task<ICollectionView> LoadMachinesAsync()
        {
            var workAreaIds = UserInfo.User.AccountWorkAreas
                .Select(x => x.WorkAreaId)
                .Distinct()
                .ToHashSet();

            var costIds = UserInfo.User.AccountCostUnits
                .Select(x => x.CostId)
                .ToHashSet();

            var swTotal = Stopwatch.StartNew();

            var sw = Stopwatch.StartNew();
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            var machines = await db.Ressources
                .AsNoTracking()
                .Include(x => x.WorkArea)
                .Include(x => x.RessourceCostUnits)
                .Where(x => workAreaIds.Contains(x.WorkAreaId ?? 0))
                .OrderBy(x => x.WorkArea.Sort)
                .ThenBy(x => x.Sort)
                .ThenBy(x => x.RessourceId)
                .ToListAsync();

            sw.Restart();
            var proc = await GetVorgangsAsync(null, machines.Select(x => x.RessourceId).ToHashSet());
            Debug.WriteLine($"GetVorgangsAsync: {sw.ElapsedMilliseconds} ms");
            _Logger.LogInformation($"GetVorgangsAsync: {sw.ElapsedMilliseconds} ms");

            sw.Restart();
            var procByRid = proc
                .Where(x => x.Rid.HasValue)
                .GroupBy(x => x.Rid.Value)
                .ToDictionary(g => g.Key, g => g.First());
            Debug.WriteLine($"procByRid: {sw.ElapsedMilliseconds} ms");
            _Logger.LogInformation($"procByRid: {sw.ElapsedMilliseconds} ms");
            sw.Restart();
            var shiftCalendars = await db.ShiftCalendars
                .AsNoTracking()
                .OrderBy(x => x.Id)
                .ToDictionaryAsync(x => x.Id, x => x.CalendarName);

            var resourceIds = machines.Select(x => x.RessourceId).ToHashSet();
            var shiftCalendarIds = machines
                .Where(x => x.ShiftCalendar.HasValue)
                .Select(x => x.ShiftCalendar!.Value)
                .Distinct()
                .ToHashSet();
            var shiftPlanCalendars = await db.ShiftCalendars
                .AsNoTracking()
                .Where(x => shiftCalendarIds.Contains(x.Id))
                .Include(x => x.ShiftCalendarShiftPlans)
                .ThenInclude(x => x.Plan)
                .AsSplitQuery()
                .ToDictionaryAsync(x => x.Id);
            var today = DateTime.Today;
            var stoppagesByResource = (await db.Stopages
                    .AsNoTracking()
                    .Where(x => resourceIds.Contains(x.Rid) && x.Endtime > today)
                    .OrderBy(x => x.Starttime)
                    .ToListAsync())
                .GroupBy(x => x.Rid)
                .ToDictionary(g => g.Key, g => g.ToList());

            var factory = _container.Resolve<PlanMachineFactory>();
            var result = new List<PlanMachine>(machines.Count);
            Debug.WriteLine($"DB-Abfrage Maschinen: {sw.ElapsedMilliseconds} ms");
            _Logger.LogInformation($"DB-Abfrage Maschinen: {sw.ElapsedMilliseconds} ms");

            sw.Restart();
            foreach (var m in machines)
            {
                if (procByRid.TryGetValue(m.RessourceId, out var assigned))
                    result.Add(factory.CreatePlanMachine(assigned.RidNavigation, _DbCtx, shiftCalendars, stoppagesByResource, shiftPlanCalendars));
                else
                    result.Add(factory.CreatePlanMachine(m, _DbCtx, shiftCalendars, stoppagesByResource, shiftPlanCalendars));
            }
            Debug.WriteLine($"PlanMachineFactory: {sw.ElapsedMilliseconds} ms");
            _Logger.LogInformation($"PlanMachineFactory: {sw.ElapsedMilliseconds} ms");

            Debug.WriteLine($"GESAMT: {swTotal.ElapsedMilliseconds} ms");
            _Logger.LogInformation($"GESAMT: {swTotal.ElapsedMilliseconds} ms");

            _machines.Clear();
            _machines.AddRange(result);

            var processByMachine = proc
                .Where(x =>
                    x.ArbPlSapNavigation != null &&
                    x.ArbPlSapNavigation.RessourceId.HasValue &&
                    costIds.Contains(x.ArbPlSapNavigation.CostId ?? 0))
                .GroupBy(x => x.ArbPlSapNavigation!.RessourceId!.Value)
                .ToDictionary(g => g.Key, g => g.ToList());

            var list = new List<Vorgang>();
            foreach (var machine in result)
            {
                if (processByMachine.TryGetValue(machine.Rid, out var items))
                    list.AddRange(items);
            }

            Priv_processes = list
                .Where(x => x.Rid == null)
                .OrderBy(x => x.Vnr)
                .ToObservableCollection();

            Priv_parking = list
                .Where(x => x.Rid < 0)
                .ToObservableCollection();

            ProcessViewSource.Source = Priv_processes;
            ParkingViewSource.Source = Priv_parking;

            RessCV = CollectionViewSource.GetDefaultView(_machines);

            if (_machines.Count > 0)
            {
                var defaultWorkAreaId = WorkAreas.FirstOrDefault()?.WorkAreaId ?? _machines[0].WorkArea?.WorkAreaId ?? 0;
                RessCV.MoveCurrentTo(_machines.FirstOrDefault(x => x.WorkArea?.WorkAreaId == defaultWorkAreaId) ?? _machines[0]);
                _currentWorkArea = ((PlanMachine)RessCV.CurrentItem)?.WorkArea?.WorkAreaId ?? 0;
            }

            NotifyPropertyChanged(() => ProcessCV);
            NotifyPropertyChanged(() => ParkingCV);

            RessCV.Filter = f => (f as PlanMachine)?.WorkArea?.WorkAreaId == _currentWorkArea &&
                                 (f as PlanMachine)!.Vis;

            ParkingCV.Filter = f => (f as Vorgang)?.Rid == _currentWorkArea * -1;
            ProcessViewSource.Filter += ProcessCV_Filter;

            return RessCV;
        }

        private async Task SubscribeEventsAsync()
        {
            await Task.Run(() =>
            {
                if (MachineTask.IsSuccessfullyCompleted)
                {
                    _ea.GetEvent<MessageOrderChanged>().Subscribe(MessageOrderReceived);
                    _ea.GetEvent<MessageVorgangChanged>().Subscribe(MessageVorgangReceived);
                    _ea.GetEvent<EnableAutoSave>().Subscribe(AutoSaveEnable);
                    _ea.GetEvent<MessagePlanmachineProcessRemoved>().Subscribe(MessagePlug);
                    return;
                }
            });
        }
        private void SelectionChange(object commandParameter)
        {
            int work = 0;
            try
            {
                if (commandParameter != null)
                {
                    if (commandParameter is SelectionChangedEventArgs sel)
                    {
                        if (sel.AddedItems[0] is WorkArea wa)
                        {
                            work = wa.WorkAreaId;
                            _currentWorkArea = wa.WorkAreaId;
                            ProcessCV?.Refresh();
                            ParkingCV?.Refresh();
                            RessCV?.Refresh();
                        }
                    } 
                }
            }

            catch (Exception e)
            {
                _Logger.LogError("{message} CommandParameter: {par} WorkArea: {work} ProcessCV {pro} Parking {perk} RessCV {res}",
                    e.ToString(), commandParameter.ToString(), work.ToString(), ProcessCV, ParkingCV, RessCV);
            }
        }
        private void OnTextSearch(object commandParameter)
        {
            try
            {
                if (commandParameter is string change)
                {
                    if (change.Length > 3 || string.IsNullOrEmpty(change))
                    {
                        _searchFilterText = change;
                        ProcessCV.Refresh();
                        _ea.GetEvent<SearchTextFilter>().Publish(_searchFilterText);
                    }
                }
                
            }
            catch (Exception e)
            {
                _Logger.LogError("{message} commandParameter= {1} ProcessCV {proc}", e.ToString(), commandParameter, ProcessCV);
            }
        }

        private void ProcessCV_Filter(object sender, FilterEventArgs e)
        {
            try
            {
                Vorgang v = (Vorgang)e.Item;
                e.Accepted = false;
                var l = _machines.Where(x => x.WorkArea?.WorkAreaId == _currentWorkArea);
               if (l.Any(x => x.Rid == v.ArbPlSapNavigation?.RessourceId))
                {
                    e.Accepted = true;
                    if (!string.IsNullOrWhiteSpace(_searchFilterText))
                    {
                        if (!(e.Accepted = v.Aid.Contains(_searchFilterText, StringComparison.CurrentCultureIgnoreCase)))
                            if (!(e.Accepted = v.AidNavigation?.Material?.Contains(_searchFilterText, StringComparison.CurrentCultureIgnoreCase) ?? false))
                                e.Accepted = v.AidNavigation?.MaterialNavigation?.Bezeichng?.Contains(_searchFilterText, StringComparison.CurrentCultureIgnoreCase) ?? false;
                    }
                    if (SearchArbpl != string.Empty)
                        e.Accepted = v.ArbPlSap?.Contains(SearchArbpl, StringComparison.CurrentCultureIgnoreCase) ?? false;
                }
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "ProcessCV Filter", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public void DragOver(IDropInfo dropInfo)
        {           
            if (PermissionsProvider.GetInstance().GetRelativeUserPermission(Permissions.MachDrop, _currentWorkArea))
            {
                if (dropInfo.Data is Vorgang)
                {
                    dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                    dropInfo.Effects = DragDropEffects.Move;
                }
            }
            if (PermissionsProvider.GetInstance().GetRelativeUserPermission(Permissions.MachSort, _currentWorkArea))
            {
                if(dropInfo.Data is PlanMachine)
                {
                    dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                    dropInfo.Effects = DragDropEffects.Move;
                }
            }
        }

        public void Drop(IDropInfo dropInfo)

        {
            string? name = dropInfo.VisualTarget.GetValue(FrameworkElement.NameProperty) as string;
            try
            {
                if (dropInfo.Data is Vorgang vrg)
                {
                    if (name == "parking")
                    {
                        int parkRid = _currentWorkArea * -1;

                        if (_DbCtx.Ressources.All(x => x.RessourceId != parkRid))
                        {
                            using (_lock.EnterScope())
                            {
                                _DbCtx.Database.ExecuteSqlRaw(@"SET IDENTITY_INSERT dbo.Ressource ON");
                                _DbCtx.Database.ExecuteSqlRaw(@"INSERT INTO dbo.Ressource(RessourceId) VALUES({0})", parkRid);
                                _DbCtx.Database.ExecuteSqlRaw(@"SET IDENTITY_INSERT dbo.Ressource OFF");
                            }
                            _Logger.LogInformation("{message} {park} {id}", "Drop Parking", parkRid, vrg.VorgangId);
                        }
                        vrg.Rid = parkRid;
                        vrg.SortPos = null;
                        _DbCtx.Update(vrg);
                    }
                    else
                    {
                        _Logger.LogInformation("{message} {0} unDrop machine", vrg.VorgangId, vrg.Rid);
                        vrg.Rid = null;
                        vrg.SortPos = null;
                        _DbCtx.Update(vrg);
                    }
                    var source = ((ListCollectionView)dropInfo.DragInfo.SourceCollection);
                    if (source.IsAddingNew) { source.CommitNew(); }
                    source.Remove(vrg);
                    ((ListCollectionView)dropInfo.TargetCollection).AddNewItem(vrg);
                    ((ListCollectionView)dropInfo.TargetCollection).CommitNew();
                }
                else if (dropInfo.Data is PlanMachine plm)
                {
                    var s = dropInfo.DragInfo.SourceCollection as ListCollectionView;
                    var t = dropInfo.TargetCollection as ListCollectionView;
                    if (s.CanRemove) s.Remove(plm);
                    var v = dropInfo.InsertIndex;

                    Debug.Assert(t != null, nameof(t) + " != null");
                    ((IList)t.SourceCollection).Insert(v, plm);

                    using (_lock.EnterScope())
                    {
                        int i = 0;
                        foreach (var m in t.SourceCollection)
                        {
                            
                            var mm = (PlanMachine)m;
                            var vv = _DbCtx.Ressources.First(x => x.RessourceId == mm.Rid );
                            vv.Sort = vv.Visability ? i : 1000;
                            _DbCtx.Update(vv);
                            i++;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
                MessageBox.Show(e.Message, "Method Drop", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private bool Changed()
        {
            using (_lock.EnterScope())
            {
                return _DbCtx.ChangeTracker.HasChanges();
            }
        }
 
        public void Closing()
        {
            try
            {
                foreach (var m in _machines)
                {
                    IViewModel vm = (IViewModel)m;
                    vm.Closing();
                }
                using (_lock.EnterScope())
                {
                    if (_DbCtx.ChangeTracker.HasChanges())
                    {
                        if (_settingsService.IsSaveMessage)
                        {
                            var result = MessageBox.Show("Sollen die Änderungen in Teamleiter-Zuteilungen gespeichert werden?",
                                Title, MessageBoxButton.YesNo, MessageBoxImage.Question);
                            if (result == MessageBoxResult.Yes)
                            {
                                _DbCtx.SaveChanges();
                            }
                        }
                        else _DbCtx.SaveChanges();
                    }
                }
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
                
            }
        }

        private class ArrayKeyComparer : IComparer<int[]>
        {
            public int Compare(int[]? x, int[]? y)
            {
                if (x[0].CompareTo(y[0])==0)
                {
                    return x[1].CompareTo(y[1]);
                }
                return x[0].CompareTo(y[0]);
            }
        }
    }
}
