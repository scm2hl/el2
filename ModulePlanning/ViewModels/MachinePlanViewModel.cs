
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
            try
            { 
                foreach (var item in list)
                {
                    if ((item != null))
                    {
                        var vo = Priv_processes?.FirstOrDefault(x => x.VorgangId == item.Value.Item2);

                        if (vo != null)
                        {
                            _ = Task.Factory.StartNew(() =>
                            {
                                if (_DbCtx.Entry(vo).State == EntityState.Modified) _DbCtx.SaveChanges();
                                _DbCtx.Entry(vo).Reload();
                                vo.RunPropertyChanged();
                                _Logger.LogInformation("reloaded {message}-{0} rid {1} {2}", vo.Aid, vo.Vnr, vo.Rid, item.Value.Item1);
                            });

                            if (item.Value.Item1 == "EL2")
                            {
                                if (vo.Rid != null)
                                {
                                    Application.Current.Dispatcher.InvokeAsync(() => Priv_processes?.Remove(vo));
                                    
                                    _Logger.LogInformation("user has added to machine {message}-{0} rid {1}", vo.Aid, vo.Vnr, vo.Rid);
                                }
 
                            }
                        }
                    }
                }                
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                //MessageBox.Show(ex.Message, "MessageVorgangReceved MachinePlan", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MessageOrderReceived(List<(string, string)?> list)
        {
            Vorgang vrg;
            Vorgang vo;
            try
            {
 
                foreach (var item in list)
                {
                    if (item == null) continue;
                    if (Priv_processes?.Any(x => x.Aid == item.Value.Item2) ?? false)
                    {
                        _ = Task.Factory.StartNew(async () =>
                        {
                            var proc = await GetVorgangsAsync(item.Value.Item2);
                            foreach (var item2 in proc)
                            {
                                if (IsRelevant(item2.ArbPlSap) && item2.Rid == null)
                                {
                                    if (Priv_processes.Contains(item2))
                                    {
                                        foreach (var item3 in item2.AidNavigation.OrderComponents)
                                        {
                                            _ = Task.Factory.StartNew(() =>
                                            {
                                                if (_DbCtx.Entry(item3).State == EntityState.Modified) { _DbCtx.SaveChanges(); }
                                                 _DbCtx.Entry(item3).Reload();
                                                
                                            });
                                        }
                                    }
                                    else
                                    {
                                        _Logger.LogInformation("scope pool added {message}-{0}", item2.Aid, item2.Vnr);
                                        using (_lock.EnterScope())
                                        {
                                            _ = Application.Current.Dispatcher.InvokeAsync(() => Priv_processes?.Add(item2));
                                            _DbCtx.ChangeTracker.Entries<Vorgang>().First(x => x.Entity.VorgangId == item2.VorgangId).State = EntityState.Unchanged;

                                            _Logger.LogInformation("pool added {message}-{0}", item2.Aid, item2.Vnr);
                                        }
                                    }
                                }
                            }
                        });
                    }
                }
                
            }
            catch (Exception ex)
            {
                _Logger.LogError("{message}", ex.ToString());
                MessageBox.Show(ex.Message, "MessageOrderReceived MachinePlan", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
        private async Task<List<Vorgang>> GetVorgangsAsync(string? aid)
        {
            List<Vorgang> query;
    
            if (aid != null)
            {
 
                query = await _DbCtx.Vorgangs
                .Include(x => x.AidNavigation)
                .ThenInclude(x => x.OrderComponents)
                .ThenInclude(x => x.MaterialNavigation)
                .Include(x => x.AidNavigation.DummyMatNavigation)
                .Include(x => x.AidNavigation.MaterialNavigation)
                .Include(x => x.ArbPlSapNavigation)
                .Include(x => x.Responses)
                .Include(x => x.RidNavigation.WorkArea)
                .Include(x => x.RidNavigation)
                .ThenInclude(x => x.RessourceWorkshifts)
                .ThenInclude(x => x.SidNavigation)
                .Where(y => y.Aid == aid
                    && y.SysStatus != null
                    && y.Text != null
                    && y.ArbPlSapNavigation.Ressource.WorkAreaId != 5
                    && y.Text.ToLower().Contains("auftrag starten") == false
                    && y.SysStatus.Contains("RÜCK") == false)
                .ToListAsync();
                
            }
            else
            {
 
                query = await _DbCtx.Vorgangs
                .Include(x => x.AidNavigation)
                .ThenInclude(x => x.OrderComponents)
                .ThenInclude(x => x.MaterialNavigation)
                .Include(x => x.AidNavigation.DummyMatNavigation)
                .Include(x => x.AidNavigation.MaterialNavigation)
                .Include(x => x.ArbPlSapNavigation)
                .Include(x => x.Responses)
                .Include(x => x.RidNavigation.WorkArea)
                .Include(x => x.RidNavigation)
                .ThenInclude(x => x.RessourceWorkshifts)
                .ThenInclude(x => x.SidNavigation)
                .Where(y => y.AidNavigation.Abgeschlossen == false
                    && y.SysStatus != null
                    && y.Text != null
                    && y.ArbPlSapNavigation.Ressource.WorkAreaId != 5
                    && y.Text.ToLower().Contains("auftrag starten") == false
                    && y.SysStatus.Contains("RÜCK") == false)
                .AsSplitQuery().ToListAsync();
                
            }
            return query;
        }

        private void SetAutoSaveTimer()
        {
            _autoSaveTimer = new System.Timers.Timer(15000);
            _autoSaveTimer.Elapsed += OnAutoSaveAsync;
            _autoSaveTimer.AutoReset = true;
            _autoSaveTimer.Enabled = true;
        }

        private async void OnAutoSaveAsync(object? sender, ElapsedEventArgs e)
        {
            try
            {
                if (MachineTask != null && MachineTask.IsSuccessfullyCompleted)
                {

                    if (_lock.TryEnter())
                    {
                        if (_DbCtx.ChangeTracker.HasChanges()) Task.Run(async () => _DbCtx.SaveChanges());
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
            var uiContext = TaskScheduler.FromCurrentSynchronizationContext();

            var proc = await GetVorgangsAsync(null);
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var mach = await db.Ressources
                .Include(x => x.WorkArea)
                .Include(x => x.RessourceCostUnits)
                .ToListAsync();
            var mach2 = new List<Ressource>();

            foreach (var uc in UserInfo.User.AccountWorkAreas.OrderBy(x => x.WorkArea.Sort))
            {
                mach2.AddRange(mach.Where(x => uc.WorkAreaId == x.WorkAreaId));
            }

            await Task.Factory.StartNew(() =>
            {
                SortedDictionary<int[], PlanMachine> result = new SortedDictionary<int[], PlanMachine>(new ArrayKeyComparer());

                PlanMachineFactory factory = _container.Resolve<PlanMachineFactory>();
                    
                foreach (var m in mach2)
                {
                    var v = proc.FirstOrDefault(x => x.Rid == m.RessourceId);
                    if (v != null) //is in my costunit allocated
                    {
                        int[] kay = new int[2];
                        kay[0] = v.RidNavigation.Sort ?? 0;
                        kay[1] = v.Rid ?? 0;                           
                        result.TryAdd(kay, factory.CreatePlanMachine(v.RidNavigation, _DbCtx));
                    }
                    else //is not in my costunit
                    {
                        int[] kay = new int[2];
                        kay[0] = m.Sort ?? 0;
                        kay[1] = m.RessourceId;
                        result.TryAdd(kay, factory.CreatePlanMachine(m, _DbCtx));
                    }
                }

                _machines.AddRange(result.Values);

                List<Vorgang> list = new();
                foreach (PlanMachine m in _machines)
                {
                    foreach (var c in UserInfo.User.AccountCostUnits)
                    {
                        list.AddRange(proc.FindAll(x => x.ArbPlSapNavigation?.RessourceId == m.Rid &&
                            x.ArbPlSapNavigation.CostId == c.CostId));
                    }
                }
                Priv_processes = list.FindAll(x => x.Rid == null).OrderBy(x => x.Vnr)
                    .ToObservableCollection();
                Priv_parking = list.FindAll(x => x.Rid < 0)
                    .ToObservableCollection();
                ProcessViewSource.Source = Priv_processes;
                ParkingViewSource.Source = Priv_parking;

                RessCV = CollectionViewSource.GetDefaultView(_machines);
                RessCV.MoveCurrentTo(_machines.First(x => x.WorkArea?.WorkAreaId == WorkAreas[0].WorkAreaId));
                _currentWorkArea = ((PlanMachine)RessCV.CurrentItem).WorkArea?.WorkAreaId ?? 0;

            }, CancellationToken.None, TaskCreationOptions.LongRunning, uiContext);
      
            NotifyPropertyChanged(() => ProcessCV);
            NotifyPropertyChanged(() => ParkingCV);
            RessCV?.Filter = f => (f as PlanMachine)?.WorkArea?.WorkAreaId == _currentWorkArea &&
                        (f as PlanMachine).Vis;
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
                    }
                    else
                    {
                        _Logger.LogInformation("{message} {0} unDrop machine", vrg.VorgangId, vrg.Rid);
                        vrg.Rid = null;
                        vrg.SortPos = null;
                        var vdb = _DbCtx.Vorgangs.First(x => x.VorgangId == vrg.VorgangId);
                        vdb.Rid = null;
                        vdb.SortPos = null;                        
                    }
                    var source = ((ListCollectionView)dropInfo.DragInfo.SourceCollection);
                    if (source.IsAddingNew) { source.CommitNew(); }
                    source.Remove(vrg);
                    ((ListCollectionView)dropInfo.TargetCollection).AddNewItem(vrg);
                    ((ListCollectionView)dropInfo.TargetCollection).CommitNew();
                }
                else if(dropInfo.Data is PlanMachine plm)
                {
                    var s = dropInfo.DragInfo.SourceCollection as ListCollectionView;
                    var t = dropInfo.TargetCollection as ListCollectionView;
                    if (s.CanRemove) s.Remove(plm);
                    var v = dropInfo.InsertIndex;

                        Debug.Assert(t != null, nameof(t) + " != null");
                        ((IList)t.SourceCollection).Insert(v, plm);
                }
                using (_lock.EnterScope())
                {
                    for (var i = 0; i < _machines.Count; i++)
                    {
                        var vv = _DbCtx.Ressources.First(x => x.RessourceId == _machines[i].Rid);
                        vv.Sort = vv.Visability ? i : 1000;
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
