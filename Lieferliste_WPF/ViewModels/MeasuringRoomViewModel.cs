using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using GongSolutions.Wpf.DragDrop;
using Lieferliste_WPF.Planning;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace Lieferliste_WPF.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class MeasuringRoomViewModel : ViewModelBase, IDropTarget, IViewModel
    {
        public string Title { get; } = "Messraum Zuteilung";
        public bool HasChange => _dbctx.ChangeTracker.HasChanges();
        private IContainerProvider _container;
        private IApplicationCommands _applicationCommands;
        private IUserSettingsService _settingsService;
        ILogger _logger;
        private RelayCommand? _textChangedCommand;
        public ICommand TextChangedCommand => _textChangedCommand ??= new RelayCommand(OnTextChanged);
        public ICommand? SaveCommand { get; private set; }
        public NotifyTaskCompletion<ICollectionView> MemberTask { get; private set; }
        public NotifyTaskCompletion<ICollectionView> VorgangTask { get; private set; }
        private DB_COS_LIEFERLISTE_SQLContext _dbctx;
        private ObservableCollection<Vorgang> _vorgangsList = new();
        private ObservableCollection<PlanWorker> _emploeeList = new();
        private ObservableCollection<Vorgang> _msfList = new();
        private string _searchText = string.Empty;
        private static System.Timers.Timer? _autoSaveTimer;
        public double SizePercent { get; }
        public ICollectionView EmploeeList { get; private set; }
        public ICollectionView VorgangsView { get; private set; }
        public ICollectionView MsfList { get; private set; }

        public MeasuringRoomViewModel(IContainerProvider container, IApplicationCommands applicationCommands, IUserSettingsService settingsService)
        {
            _container = container;
            _applicationCommands = applicationCommands;
            _settingsService = settingsService;
            _dbctx = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var factory = _container.Resolve<ILoggerFactory>();
            _logger = factory.CreateLogger<MeasuringRoomViewModel>();
            SizePercent = 340 * _settingsService.SizePercent / 100;
            MemberTask = new NotifyTaskCompletion<ICollectionView>(LoadMemberAsync());
            SaveCommand = new ActionCommand(OnSaveExecuted, OnSaveCanExecute);
            VorgangsView = CollectionViewSource.GetDefaultView(_vorgangsList);
            VorgangsView.Filter += FilterPredicate;
            MsfList = CollectionViewSource.GetDefaultView( _msfList);
            if (_settingsService.IsAutoSave) SetAutoSave();
        }

        private bool OnSaveCanExecute(object arg)
        {
            if (_settingsService.IsAutoSave) return false;
            return _dbctx.ChangeTracker.HasChanges();
        }

        private void OnSaveExecuted(object obj)
        {
            _dbctx.SaveChanges();
        }
        private void SetAutoSave()
        {
            _autoSaveTimer = new System.Timers.Timer(60000);
            _autoSaveTimer.Elapsed += OnAutoSave;
            _autoSaveTimer.AutoReset = true;
            _autoSaveTimer.Enabled = true;
        }

        private void OnAutoSave(object? sender, ElapsedEventArgs e)
        {
            if (_dbctx.ChangeTracker.HasChanges()) _dbctx.SaveChanges();
        }

        private async Task<ICollectionView> LoadDataAsync()
        {
            try
            {
                var ord = await _dbctx.Vorgangs
            .Include(x => x.AidNavigation)
            .ThenInclude(x => x.MaterialNavigation)
            .Include(x => x.AidNavigation.DummyMatNavigation)
            .Include(x => x.ArbPlSapNavigation)
            .Include(x => x.MeasureRessVorgangs)
            .Where(x => (x.ArbPlSapNavigation.Ressource != null)
                && x.AidNavigation.Abgeschlossen == false
                && ((x.SysStatus != null) && x.SysStatus.Contains("RÜCK") == false))
            .ToListAsync();

                _vorgangsList.AddRange(ord.ExceptBy(_dbctx.MeasureRessVorgangs.Select(x => x.VorgId), x => x.VorgangId));
                VorgangsView = CollectionViewSource.GetDefaultView(_vorgangsList);
                VorgangsView.Filter += FilterPredicate;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "MeasuringRoom Load Data", MessageBoxButton.OK, MessageBoxImage.Error);
                _logger.LogError("{message}", ex.ToString());
            }
            return VorgangsView;
        }

        private async Task<ICollectionView> LoadMemberAsync()
        {
            try
            {
                var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var mem = await db.MeasureResses
                  .ToListAsync();
                var ord = await _dbctx.Vorgangs
                    .Include(x => x.AidNavigation)
                    .ThenInclude(x => x.MaterialNavigation)
                    .Include(x => x.AidNavigation.DummyMatNavigation)
                    .Include(x => x.ArbPlSapNavigation)
                    .Include(x => x.MeasureRessVorgangs)
                    .Where(x => (x.MeasureRessVorgangs != null)
                    && x.ArbPlSapNavigation.Ressource.MeasureRessRessources.Any()
                    && x.AidNavigation.Abgeschlossen == false
                    && ((x.SysStatus != null) && x.SysStatus.Contains("RÜCK") == false))
                    .ToListAsync();
                var factory = _container.Resolve<PlanWorkerFactory>();
                foreach (var item in mem)
                {
                    var vorg = ord.Where(x => x.MeasureRessVorgangs.Any(x => x.MessId == item.MessRid)).ToList();
                    _emploeeList.Add(factory.CreatePlanWorker(item.MessRid, vorg));
                }

                _vorgangsList.AddRange(ord.ExceptBy(_dbctx.MeasureRessVorgangs.Select(x => x.VorgId), x => x.VorgangId));
                EmploeeList = CollectionViewSource.GetDefaultView(_emploeeList);
                _msfList.AddRange(ord.Where(x => x.MsfInWork));
                return EmploeeList;
            }
            catch (Exception e)
            {
                _logger.LogError("{message}", e.ToString());
                return EmploeeList;
            }

        }
        private void OnTextChanged(object obj)
        {
            if (obj is string s)
            {
                _searchText = s;
                VorgangsView.Refresh();
            }
        }
        private bool FilterPredicate(object obj)
        {
            bool accepted = true;
            if (obj is Vorgang vrg && _searchText != string.Empty)
            {
                if (!(accepted = vrg.Aid.Contains(_searchText, StringComparison.CurrentCultureIgnoreCase)))
                    if (!(accepted = vrg.AidNavigation.Material?.Contains(_searchText, StringComparison.CurrentCultureIgnoreCase) ?? false))
                        accepted = vrg.AidNavigation.MaterialNavigation?.Bezeichng?.Contains(_searchText, StringComparison.CurrentCultureIgnoreCase) ?? false;

            }
            return accepted;
        }
        public void DragOver(IDropInfo dropInfo)
        {
            if (PermissionsProvider.GetInstance().GetUserPermission(Permissions.MessDrop))
            {
                if (dropInfo.Data is Vorgang)
                {
                    dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                    dropInfo.Effects = DragDropEffects.Move;
                }
            }
        }

        public void Drop(IDropInfo dropInfo)
        {
            try
            {
                if (dropInfo.Data is Vorgang vrg)
                {
                    var source = ((ListCollectionView)dropInfo.DragInfo.SourceCollection);
                    if (source.IsAddingNew) { source.CommitNew(); }
                    source.Remove(vrg);
                    ((ListCollectionView)dropInfo.TargetCollection).AddNewItem(vrg);
                    ((ListCollectionView)dropInfo.TargetCollection).CommitNew();
                    var t = dropInfo.VisualTarget.GetType().GetProperty("Name");
                    
                    if ("MSF".Equals(t?.GetValue(dropInfo.VisualTarget, null)))
                    {
                        vrg.MsfInWork = true;
                    }
                    else
                    {   vrg.MsfInWork = false;
                        _dbctx.MeasureRessVorgangs.RemoveRange(_dbctx.MeasureRessVorgangs.Where(x => x.VorgId.Trim() == vrg.VorgangId));
                        _logger.LogInformation("drops {message}", vrg.VorgangId);
                    }
                }
            }
            catch (Exception e)
            {
                _logger.LogError("{message}", e.ToString());
            }
        }

        public void Closing()
        {
            if (_dbctx.ChangeTracker.HasChanges())
            {
                if (_settingsService.IsSaveMessage)
                {
                    var result = MessageBox.Show(string.Format("Sollen die Änderungen in {0} gespeichert werden?", Title),
                        Title, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        _dbctx.SaveChanges();
                    }
                }
                else _dbctx.SaveChanges();
            }
        }
    }
}
