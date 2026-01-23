using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Lieferliste_WPF.Utilities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using ModulePlanning.Planning;
using Prism.Dialogs;
using Prism.Events;
using Prism.Ioc;
using Prism.Navigation;
using Prism.Navigation.Regions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Printing;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Lieferliste_WPF.ViewModels
{
    /// <summary>
    /// Class for the main window's view-model.
    /// </summary>
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public class MainWindowViewModel : ViewModelBase
    {
        public ICommand OpenMachinePlanCommand { get; private set; }
        public ICommand OpenLieferlisteCommand { get; private set; }
        public ICommand OpenSettingsCommand { get; private set; }
        public ICommand OpenUserMgmtCommand { get; private set; }
        public ICommand OpenRoleMgmtCommand { get; private set; }
        public ICommand OpenMachineMgmtCommand { get; private set; }
        public ICommand TabCloseCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }
        public ICommand OpenSeclusionsCommand { get; private set; }
        public ICommand OpenWorkAreaCommand { get; private set; }
        public ICommand OpenProjectCombineCommand { get; private set; }
        public ICommand OpenMeasuringCommand { get; private set; }
        public ICommand OpenShiftCommand { get; private set; }
        public ICommand OpenHolidayCommand { get; private set; }
        public ICommand FollowMsfCommand { get; private set; }
        public ICommand OpenReportCommand { get; private set; }
        public ICommand OpenNoteCommand { get; private set; }

        private IApplicationCommands _applicationCommands;
        public IApplicationCommands ApplicationCommands
        {
            get { return _applicationCommands; }
            set
            {
                if (_applicationCommands != value)
                {
                    _applicationCommands = value;
                    NotifyPropertyChanged(() => ApplicationCommands);
                }
            }
        }
        public ICommand OpenMeasureOperCommand { get; }
        public ICommand ExplorerCommand { get; }
        public ICommand OpenOrderCommand { get; }
        public ICommand OrderCloseCommand { get; }
        public ICommand OpenProjectOverViewCommand { get; }
        public ICommand MachinePrintCommand { get; }
        public ICommand OpenProductViewCommand { get; }
        private NotifyTaskCompletion<int>? _onlineTask;
        public NotifyTaskCompletion<int>? OnlineTask
        {
            get { return _onlineTask; }
            set
            {
                if (_onlineTask != value)
                {
                    _onlineTask = value;
                    NotifyPropertyChanged(() => OnlineTask);
                }
            }
        }
        private string? _msg;
        public string? Msg
        {
            get { return _msg; }
            private set
            {
                if (value != _msg)
                {
                    _msg = value;
                    NotifyPropertyChanged(() => Msg);
                }
            }
        }
        private WorkareaDocumentInfo? _workareaDocumentInfo;
        private static int _onlines;
        private static TimerCallback? timerCallbackOnl;
        private static TimerCallback? timerCallbackMsg;
        private readonly Timer _timerOnl;
        private readonly Timer _timerMsg;

        private IRegionManager _regionmanager;
        private static readonly Lock _lock = new();
        private readonly IContainerExtension _container;
        private readonly IDialogService _dialogService;
        private readonly IEventAggregator _ea;
        private readonly IUserSettingsService _settingsService;
        private readonly ILogger _Logger;
        public MainWindowViewModel(IRegionManager regionManager,
            IContainerExtension container,
            IApplicationCommands applicationCommands,
            IDialogService dialogService,
            IEventAggregator ea,
            IUserSettingsService settingsService)
        {

            try
            {
                _regionmanager = regionManager;
                _container = container;
                _applicationCommands = applicationCommands;
                _dialogService = dialogService;
                _ea = ea;
                _settingsService = settingsService;
                var loggerFactory = _container.Resolve<Microsoft.Extensions.Logging.ILoggerFactory>();

                _Logger = loggerFactory.CreateLogger<MainWindowViewModel>();

                if (CoreFunction.PriorProcess == null)
                {
                    _timerMsg.Dispose();
                    _timerOnl.Dispose();
                    App.Current.Shutdown();
                }
                _ = RegisterMe();
                timerCallbackOnl = new(ExecuteT);
                timerCallbackMsg = new(ExecuteGetMsg);
                _timerOnl = new Timer(callback: timerCallbackOnl, null, 0, 59000);
                _timerMsg = new Timer(callback: timerCallbackMsg, null, 0, 60000);

                TabCloseCommand = new ActionCommand(OnTabCloseExecuted, OnTabCloseCanExecute);
                CloseCommand = new ActionCommand(OnCloseExecuted, OnCloseCanExecute);
                _applicationCommands.CloseCommand.RegisterCommand(CloseCommand);
                ExplorerCommand = new ActionCommand(OnOpenExplorerExecuted, OnOpenExplorerCanExecute);
                _applicationCommands.ExplorerCommand.RegisterCommand(ExplorerCommand);
                OpenOrderCommand = new ActionCommand(OnOpenOrderExecuted, OnOpenOrderCanExecute);
                _applicationCommands.OpenOrderCommand.RegisterCommand(OpenOrderCommand);
                OrderCloseCommand = new ActionCommand(OnOrderCloseExecuted, OnOrderCloseCanExecute);
                _applicationCommands.OrderCloseCommand.RegisterCommand(OrderCloseCommand);
                OpenProjectOverViewCommand = new ActionCommand(OnOpenProjectOverViewExecuted, OnOpenProjectOverViewCanExecute);
                _applicationCommands.OpenProjectOverViewCommand.RegisterCommand(OpenProjectOverViewCommand);
                MachinePrintCommand = new ActionCommand(OnMachinePrintExecuted, OnMachinePrintCanExecute);
                _applicationCommands.MachinePrintCommand.RegisterCommand(MachinePrintCommand);
                OpenMeasureOperCommand = new ActionCommand(OnOpenMeasureOperExecuted, OnOpenMeasureOperCanExecute);
                _applicationCommands.OpenMeasuringOperCommand.RegisterCommand(OpenMeasureOperCommand);
                FollowMsfCommand = new ActionCommand(OnFollowMsfExecuted, OnFollowMsfCanExecute);
                _applicationCommands.FollowMsfCommand.RegisterCommand(FollowMsfCommand);

                OpenLieferlisteCommand = new ActionCommand(OnOpenLieferlisteExecuted, OnOpenLieferlisteCanExecute);
                OpenMachinePlanCommand = new ActionCommand(OnOpenMachinePlanExecuted, OnOpenMachinePlanCanExecute);
                OpenUserMgmtCommand = new ActionCommand(OnOpenUserMgmtExecuted, OnOpenUserMgmtCanExecute);
                OpenRoleMgmtCommand = new ActionCommand(OnOpenRoleMgmtExecuted, OnOpenRoleMgmtCanExecute);
                OpenMachineMgmtCommand = new ActionCommand(OnOpenMachineMgmtExecuted, OnOpenMachineMgmtCanExecute);
                OpenSettingsCommand = new ActionCommand(OnOpenSettingsExecuted, OnOpenSettingsCanExecute);
                OpenSeclusionsCommand = new ActionCommand(OnOpenSeclusionsExecuted, OnOpenSeclusionsCanExecute);
                OpenWorkAreaCommand = new ActionCommand(OnOpenWorkAreaExecuted, OnOpenWorkAreaCanExecute);
                OpenMeasuringCommand = new ActionCommand(OnOpenMeasuringExecuted, OnOpenMeasuringCanExecute);
                OpenProjectCombineCommand = new ActionCommand(OnOpenProjectCombineExecuted, OnOpenProjectCombineCanExecute);
                OpenHolidayCommand = new ActionCommand(OnOpenHolidayExecuted, OnOpenHolidayCanExecute);
                OpenShiftCommand = new ActionCommand(OnOpenShiftExecuted, OnOpenShiftCanExecute);                
                OpenReportCommand = new ActionCommand(OnOpenReportExecuted, OnOpenReportCanExecute);
                OpenProductViewCommand = new ActionCommand(OnOpenProductExecuted, OnOpenProductCanExecute);
                OpenNoteCommand = new ActionCommand(OnOpenNoteExecuted, OnOpenNoteCanExecute);

                _workareaDocumentInfo = new WorkareaDocumentInfo(container);


                //DbOperations();
            }
            catch (Exception ex)
            {
                _Logger?.LogError("{message}", ex);
                MessageBox.Show(ex.ToString());
            }
        }

        private bool OnOpenNoteCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.EmployNoteOpen);
        }

        private void OnOpenNoteExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("EmployNote", UriKind.Relative));
        }

        private bool OnFollowMsfCanExecute(object arg)
        {
            return true;
        }

        private void OnFollowMsfExecuted(object obj)
        {
            if (obj is string s)
            {
                if (RuleInfo.Rules.TryGetValue("MeasureMsfDomain", out Rule? msf))
                {
                    System.Diagnostics.Process.Start("explorer", msf.RuleValue + s);
                }
            }
        }

        private bool OnOpenProductCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.ProductsViewOpen);
        }

        private void OnOpenProductExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("Products", UriKind.Relative));
        }

        private bool OnOpenReportCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.ReportOpen);
        }

        private void OnOpenReportExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("ReportMainView", UriKind.Relative));
        }

        #region Commands
        private bool OnOpenMeasureOperCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenMeaOper);
        }

        private void OnOpenMeasureOperExecuted(object obj)
        {
            var navi = new NavigationParameters();
            if (obj is Vorgang par) navi.Add("order", par);
                 
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("MeasuringDocuments", UriKind.Relative), navi);
        }
        private bool OnOpenShiftCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.ShiftEdit);
        }

        private void OnOpenShiftExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("ShiftPlanEdit", UriKind.Relative));
        }

        private bool OnOpenHolidayCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.HoliEdit);
        }

        private void OnOpenHolidayExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("HolidayEdit", UriKind.Relative));
        }

        private bool OnOpenProjectCombineCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenProjectCombine);
        }

        private void OnOpenProjectCombineExecuted(object obj)
        {
            _dialogService.Show("ProjectEdit");
        }



        private bool OnOpenProjectOverViewCanExecute(object arg)
        {
            bool accept = false;
            if (arg is Vorgang) { accept = true; }
            else if (arg is PspNode<Shape> shape) { accept = shape.HasOrder; }

            return accept && PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenProject);
        }
        private void OnOpenProjectOverViewExecuted(object obj)
        {
            string param = string.Empty;
            if (obj is Vorgang vrg)
            {
                param = vrg.AidNavigation.ProId ??= string.Empty;
            }
            else if (obj is string s) { param = s; }
            else if (obj is PspNode<Shape> shape) { param = shape.Node.ToString(); }
            if (param.IsNullOrEmpty()) { return; }

            var par = new DialogParameters();
            par.Add("projectNo", param);
            _dialogService.Show("Projects", par, null);
        }

        private bool OnOpenMeasuringCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenMeasure);
        }

        private void OnOpenMeasuringExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("MeasuringRoom", UriKind.Relative));
        }
        private bool OnOpenWorkAreaCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenWorkArea);
        }

        private void OnOpenWorkAreaExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("ShowWorkArea", UriKind.Relative));
        }

        private bool OnOpenSeclusionsCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.Seclusions);
        }

        private void OnOpenSeclusionsExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("Seclusions", UriKind.Relative));
        }
        private void OnOrderCloseExecuted(object obj)
        {
            try
            {
                if (obj is object[] onr)
                {
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    var v = db.OrderRbs.First(x => x.Aid == (string)onr[0]);
                    v.Abgeschlossen = true;
                    foreach (var item in v.Vorgangs.Where(x => x.Visability == false))
                    {
                        item.Visability = true;
                    }
                    _Logger.LogInformation("Enclosed: {message}", v.Aid);
                    db.SaveChangesAsync();
                    _ea.GetEvent<MessageOrderEnclose>().Publish(v);
                }
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
            }
        }

        private bool OnOrderCloseCanExecute(object arg)
        {
            try
            {
                if (arg is object[] onr)
                {
                    if (onr[1] is Boolean f)
                    {
                        return PermissionsProvider.GetInstance().GetUserPermission(Permissions.CloseOrder) &&
                            (f || Keyboard.IsKeyDown(Key.LeftAlt));
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
                MessageBox.Show(e.Message, "EncloseCan", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        private void OnCloseExecuted(object obj)
        {
            try
            {
                if (obj == null)
                {
                    using (var Dbctx = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>())
                    {
                        Dbctx.ChangeTracker.DetectChanges();
                        if (Dbctx.ChangeTracker.HasChanges())
                        {
                            var r = MessageBox.Show("Sollen die Änderungen noch in\n die Datenbank gespeichert werden?",
                                "MS SQL Datenbank", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
                            if (r == MessageBoxResult.Yes) Dbctx.SaveChanges();
                        }

                        var del = Dbctx.InMemoryOnlines.Where(x => UserInfo.Dbid.Equals(x.OnlId));                        
                        if (Dbctx.InMemoryMsgs.Any()) Dbctx.InMemoryMsgs.Where(x => x.OnlId.Equals(del.First().OnlId)).ExecuteDelete();
                        del.ExecuteDelete();
                    }
                }
                else
                {
                    _regionmanager.Regions[RegionNames.MainContentRegion].Remove(obj);
                }
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
            }
        }

        private bool OnCloseCanExecute(object arg)
        {
            return true;
        }

        private bool OnTabCloseCanExecute(object arg)
        {
            return true;
        }

        private void OnTabCloseExecuted(object obj)
        {
            if (obj is FrameworkElement f)
            {
                var vm = f.DataContext as IViewModel;
                vm?.Closing();
            }
            _regionmanager.Regions[RegionNames.MainContentRegion].Remove(obj);
        }
        private static bool OnOpenOrderCanExecute(object arg)
        {
            bool ret = false;
            if (arg is string f) { ret = !string.IsNullOrEmpty(f); }
            if (arg is Vorgang) ret = true;
            if (arg is Shape) ret = true;
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.Order) && ret;
                
        }
        private void OnOpenOrderExecuted(object parameter)
        {
            var aid = string.Empty;
            if (parameter is Vorgang y) aid = y.Aid;
            else if (parameter is string) aid = (string)parameter;
            else if (parameter is Shape shape) { aid = shape.ToString(); }
            if (string.IsNullOrEmpty(aid) == false)
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var vrgs = db.Vorgangs
                    .Include(x => x.AidNavigation)
                    .Include(x => x.AidNavigation.DummyMatNavigation)
                    .Include(x => x.AidNavigation.MaterialNavigation)
                    .Include(x => x.AidNavigation.Pro)
                    .Include(x => x.RidNavigation)
                    .Include(x => x.ArbPlSapNavigation)
                    .ThenInclude(x => x.Ressource)
                    .ThenInclude(x => x.WorkArea)
                    .Where(x => x.Aid == aid)
                    .ToList();

                if (vrgs != null)
                {
                    var par = new DialogParameters
                    {
                        { "vrgList", vrgs }
                    };
                    _dialogService.Show("Order", par, null);
                }
            }
        }
        private void OnOpenMachineMgmtExecuted(object selectedItem)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("MachineEdit", UriKind.Relative));
        }

        private bool OnOpenMachineMgmtCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.MachEdit);
        }

        private void OnOpenRoleMgmtExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("RoleEdit", UriKind.Relative));
        }

        private bool OnOpenRoleMgmtCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.RoleEdit);
        }

        private void OnOpenUserMgmtExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("UserEdit", UriKind.Relative));
        }

        private bool OnOpenUserMgmtCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenUserEdit);
        }

        private void OnOpenSettingsExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("UserSettings", UriKind.Relative));
        }

        private bool OnOpenSettingsCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.UserSett);
        }

        private void OnOpenMachinePlanExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("MachinePlan", UriKind.Relative));
        }

        private bool OnOpenMachinePlanCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.MachPlan);
        }

        private bool OnOpenLieferlisteCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.Liefer);
        }

        private void OnOpenLieferlisteExecuted(object obj)
        {
            _regionmanager.RequestNavigate(RegionNames.MainContentRegion, new Uri("Liefer", UriKind.RelativeOrAbsolute));
        }
        private bool OnOpenExplorerCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.OpenExpl);
        }
        private void OnOpenExplorerExecuted(object obj)
        {
            Document? docu = null;
            if (obj is null) return;
            if (obj is OrderViewModel o)
            {
                var tt = (o.Material == "DUMMY") ? o.DummyMat : o.Material;
                if (tt == null) { tt = string.Empty; }
                if (o.ArchivState > 1)
                {
                    MessageBox.Show("Für diesen Auftrag gibt es keine Dokumente!"
                        , "Info", MessageBoxButton.OK);
                    return;
                }
                else if (o.ArchivState == 1)
                {

                    if (o.ArchivPath != null)
                    {
                        if (new DirectoryInfo(o.ArchivPath).Exists)
                            Process.Start("explorer.exe", o.ArchivPath);
                        else
                            MessageBox.Show("Der Dateiordner existiert nicht", "Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    return;
                }
                docu = _workareaDocumentInfo.CreateDocumentInfos([tt, o.Aid, string.Empty], o.Material == "DUMMY");
            }
            else if (obj is Vorgang v)
            {
                var tt = string.IsNullOrEmpty(v.AidNavigation.Material) ? v.AidNavigation.DummyMat : v.AidNavigation.Material;
                if (tt == null) { tt = string.Empty; }
                docu = _workareaDocumentInfo.CreateDocumentInfos([tt, v.Aid, string.Empty], string.IsNullOrEmpty(v.AidNavigation.Material));
            }
            else if (obj is TblMaterial m)
            {
                docu = _workareaDocumentInfo.CreateDocumentInfos([m.Ttnr, string.Empty, string.Empty]);
            }
            else if (obj is OrderRb orb)
            {
                var tt = string.IsNullOrEmpty(orb.Material) ? orb.DummyMat : orb.Material;
                if (tt == null) { tt = string.Empty; }
                if (orb.ArchivState > 1)
                {
                    MessageBox.Show("Für diesen Auftrag gibt es keine Dokumente!"
                        , "Info", MessageBoxButton.OK);
                    return;
                }
                else if (orb.ArchivState == 1)
                {
                    if (orb.ArchivPath != null)
                    {
                        if (new DirectoryInfo(orb.ArchivPath).Exists)
                            Process.Start("explorer.exe", orb.ArchivPath);
                        else
                            MessageBox.Show("Der Dateiordner existiert nicht", "Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    return;
                }

                docu = _workareaDocumentInfo.CreateDocumentInfos([tt, orb.Aid, string.Empty], string.IsNullOrEmpty(orb.Material));
            }
            else if (obj is ValueTuple<string, string, int, string> dic)
            {
                if (dic.Item3 > 1)
                {
                    MessageBox.Show("Für diesen Auftrag gibt es keine Dokumente!"
                        , "Info", MessageBoxButton.OK);
                    return;
                }
                else if (dic.Item3 == 0)
                {
                    docu = _workareaDocumentInfo.CreateDocumentInfos([dic.Item1, dic.Item2, string.Empty]);
                }
                else
                {

                    Process.Start("explorer.exe", dic.Item4);
                }
            }
            else if (obj is Dictionary<string, object> dicobj)
            {

                string tt;

                tt = (string)(((string?)dicobj["ttnr"] == "DUMMY") ? dicobj["dummymat"] : dicobj["ttnr"]);

                docu = _workareaDocumentInfo.CreateDocumentInfos([tt, (string)dicobj["aid"], string.Empty], (string?)dicobj["ttnr"] == "DUMMY");
            }
            if (docu != null)
            {
                if (!Directory.Exists(docu[DocumentPart.RootPath]))
                {
                    MessageBox.Show($"Der Hauptpfad '{docu[DocumentPart.RootPath]}'\nwurde nicht gefunden!"
                        , "Error", MessageBoxButton.OK);
                }
                else
                {
                    var p = Path.Combine(docu[DocumentPart.RootPath], docu[DocumentPart.SavePath]);
                    if (Directory.Exists(p))
                        Process.Start("explorer.exe", @p);
                    else
                        MessageBox.Show(@"Der Pfad " + @p + " wurde nicht gefunden.", "Dateiexplorer", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
        }
        private bool OnMachinePrintCanExecute(object arg)
        {
            int wo = 0;
            if (arg is PlanMachine plan)
            {
                wo = plan.WorkArea?.WorkAreaId ?? 0;
            }

            if (wo != 0) { return PermissionsProvider.GetInstance().GetRelativeUserPermission(Permissions.MachPrint, wo); }
            return false;
        }

        private void OnMachinePrintExecuted(object obj)
        {
            try
            {
                var print = new PrintDialog();
                var ticket = new PrintTicket();
                ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
                ticket.PageOrientation = PageOrientation.Landscape;
                print.PrintTicket = ticket;
                PrintingProxy printingProxy = new PrintingProxy();
                if (obj is PlanMachine planMachine)
                {
                    printingProxy.PrintPreview(planMachine, ticket);
                }
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message, "MachinePrint", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion
        //private void SetTimer()
        //{
        //    // Create a timer with a 59 seconds interval.
        //    _timer = new System.Timers.Timer(59000);
        //    // Hook up the Elapsed event for the timer. 
        //    _timer.Elapsed += OnTimedEvent;
        //    _timer.AutoReset = true;
        //    _timer.Enabled = true;
        //}
        public int Onlines
        {
            get { return _onlines; }
            private set
            {
                if (_onlines != value)
                {
                    _onlines = value;
                    NotifyPropertyChanged(() => Onlines);
                }
            }
        }
        private void ExecuteT(object? state)
        {
            if (UserInfo.Dbid == 0) return;
            if (_lock.TryEnter())
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                Onlines = db.InMemoryOnlines.Count();
                if (db.InMemoryOnlines.All(x => x.OnlId != UserInfo.Dbid))
                {
                    if (MessageBox.Show(Application.Current.MainWindow,
                        "Registrierung ist abgelaufen!\nDie Anwendung wird beendet.", "Warnung", MessageBoxButton.OK, MessageBoxImage.Stop) ==
                        MessageBoxResult.OK)
                    { Application.Current.Shutdown(10); }
                }
                else
                {
                    db.Database.ExecuteSqlRaw(@"UPDATE InMemoryOnline SET LifeTime = {0} WHERE OnlID = {1}",
                    DateTime.Now,
                    UserInfo.Dbid);
                }
                _lock.Exit();
            }
        }
        //private void OnTimedEvent(object? sender, ElapsedEventArgs e)
        //{           
        //    Application.Current.Dispatcher.InvokeAsync(new Action(() =>
        //    {
        //        using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
        //        Onlines = db.InMemoryOnlines.Count();
        //        if (db.InMemoryOnlines.All(x => x.OnlId != UserInfo.Dbid))
        //        {
        //            if (MessageBox.Show(Application.Current.MainWindow,
        //                "Registrierung ist abgelaufen!\nDie Anwendung wird beendet.", "Warnung", MessageBoxButton.OK, MessageBoxImage.Stop) ==
        //                MessageBoxResult.OK)
        //            { Application.Current.Shutdown(10); }
        //        }
        //        else db.Database.ExecuteSqlRaw(@"UPDATE InMemoryOnline SET LifeTime = {0} WHERE OnlId = {1}",
        //            DateTime.Now,
        //            UserInfo.Dbid);
        //    }), System.Windows.Threading.DispatcherPriority.Background);
        //}
        private void ExecuteGetMsg(object? state)
        {


            List<(string, string)?> msgListV = [];
            List<(string, string)?> msgListO = [];

            try
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                {
                    var m = db.InMemoryMsgs.AsNoTracking()
                        .Where(x => x.OnlId == UserInfo.Dbid)
                        .OrderBy(x => x.Timestamp)
                        .ToListAsync();
                    if (m.Result.Count > 0)
                    {
                        foreach (var item in m.Result)
                        {
                            if (item != null && item.TableName == "Vorgang")
                            {
                                if (item.OldValue != item.NewValue)
                                    if (msgListV.All(x => x.Value.Item2 != item.PrimaryKey))
                                    {
                                        msgListV.Add((item.Invoker, item.PrimaryKey));
                                    }
                            }
                            if (item != null && item.TableName == "OrderRB")
                            {
                                if (item.OldValue != item.NewValue)
                                    if (msgListO.All(x => x.Value.Item2 != item.PrimaryKey))
                                    {
                                        msgListO.Add((item.Invoker, item.PrimaryKey));

                                    }
                            }
                        }

                        foreach (var msg in m.Result)
                        {
                            db.Database.ExecuteSqlRaw(@"DELETE FROM InMemoryMsg WHERE MsgId={0}", msg.MsgId);
                        }
                    }
                }
                Msg = string.Format("{0}-{1} ", DateTime.Now.ToString("HH:mm:ss"), msgListO.Count + msgListV.Count);

                if (msgListV.Count > 0)
                    _ea.GetEvent<MessageVorgangChanged>().Publish(msgListV);
                if (msgListO.Count > 0)
                    _ea.GetEvent<MessageOrderChanged>().Publish(msgListO);

            }
            catch (Exception ex)
            {
                _Logger.LogError("Auftrag:{msgo} -- Vorgang:{msgv}", [msgListO.Count, msgListV.Count]);
                _Logger.LogCritical("{message}", ex.ToString());
            }
        }
        
        //private void SetMsgDBTimer()
        //{
        //    // Create a timer with a 60 seconds interval.
        //    _timer = new System.Timers.Timer(60000);
        //    // Hook up the Elapsed event for the timer. 
        //    _timer.Elapsed += OnMsgDBTimedEvent;
        //    _timer.AutoReset = true;
        //    _timer.Enabled = true;
        //}
        //private async void OnMsgDBTimedEvent(object? sender, ElapsedEventArgs e)
        //{
        //    List<(string, string)?> msgListV = [];
        //    List<(string, string)?> msgListO = [];

        //    try
        //    {
        //        using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
        //        {
        //            var m = await db.InMemoryMsgs.AsNoTracking()
        //                .Include(x => x.Onl)
        //                .Where(x => x.Onl.OnlId == UserInfo.Dbid)
        //                .OrderBy(x => x.Timestamp)
        //                .ToListAsync();
        //            if (m.Count > 0)
        //            {
        //                foreach (var item in m)
        //                {
        //                    if (item != null && item.TableName == "Vorgang")
        //                    {
        //                        if (item.OldValue != item.NewValue)
        //                            if (msgListV.All(x => x.Value.Item2 != item.PrimaryKey))
        //                            {
        //                                msgListV.Add((item.Invoker, item.PrimaryKey));
        //                            }
        //                    }
        //                    if (item != null && item.TableName == "OrderRB")
        //                    {
        //                        if (item.OldValue != item.NewValue)
        //                            if (msgListO.All(x => x.Value.Item2 != item.PrimaryKey))
        //                            {
        //                                msgListO.Add((item.Invoker, item.PrimaryKey));
                                        
        //                            }
        //                    }                            
        //                }

        //                foreach (var msg in m)
        //                {
        //                    await db.Database.ExecuteSqlRawAsync(@"DELETE FROM InMemoryMsg WHERE MsgId={0}", msg.MsgId);
        //                }
        //            }
        //        }
        //        Msg = string.Format("{0}-{1} ", DateTime.Now.ToString("HH:mm:ss"), msgListO.Count + msgListV.Count);

        //        if (msgListV.Count > 0)
        //            _ea.GetEvent<MessageVorgangChanged>().Publish(msgListV);
        //        if (msgListO.Count > 0)
        //            _ea.GetEvent<MessageOrderChanged>().Publish(msgListO);

        //    }
        //    catch (Exception ex)
        //    {
        //        _Logger.LogError("Auftrag:{msgo} -- Vorgang:{msgv}", [msgListO.Count, msgListV.Count]);
        //        _Logger.LogCritical("{message}", ex.ToString());
        //    }
        //}

        private async System.Threading.Tasks.Task RegisterMe()
        {
            try
            {
                _lock.EnterScope();
                {
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    using var transaction = db.Database.BeginTransaction();
                    var onl = db.InMemoryOnlines.FirstOrDefault(x => x.Userid == UserInfo.User.UserId && x.PcId == UserInfo.PC);
                    if (onl != null)
                    {
                        db.Database.ExecuteSqlRaw("DELETE InMemoryMsg WHERE OnlId=@p0", onl.OnlId);
                        db.Database.ExecuteSqlRaw("DELETE InMemoryOnline WHERE OnlID=@p0", onl.OnlId);
                    }
                    db.Database.ExecuteSqlRaw(@"INSERT INTO InMemoryOnline(Userid,PcId,Login, LifeTime) VALUES({0},{1},{2}, {3})",
                        UserInfo.User.UserId,
                        UserInfo.PC ?? string.Empty,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    transaction.Commit();

                    UserInfo.Dbid = db.InMemoryOnlines.Single(x => x.Userid == UserInfo.User.UserId && x.PcId == UserInfo.PC).OnlId;
                    _Logger.LogInformation("Startup {user}-{pc}-{id}--{version}", [UserInfo.User.UserId, UserInfo.PC, UserInfo.Dbid, Assembly.GetExecutingAssembly().GetName().Version]);
                }
            }
            catch (Exception e)
            {
                _Logger.LogError("{message}", e.ToString());
            }
            finally
            {
                _lock?.Exit();
            }
        }


        private void DbOperations()
        {
            var gl = new Globals(_container);

            //List<ProjectScheme> schemes = new List<ProjectScheme>();
            //schemes.Add(new ProjectScheme("DS", "(DS-[0-9]{6})(-[0-9]{2})?(-[0-9]{2})?(-[0-9]{2})?(-[0-9]{2})?"));
            //schemes.Add(new ProjectScheme("SC-PR", "(SC-PR-[0-9]{9})(-[0-9]{2})?(-[0-9]{2})?(-[0-9]{2})?"));
            //schemes.Add(new ProjectScheme("BM", "(BM-[0-9]{8})(_[0-9]{3})(_[0-9]{8})?"));

            //gl.SaveProjectSchemes(schemes);
            var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var orders = db.OrderRbs
                .Where(x => x.ArchivState == 1 && x.ArchivPath.Length != 0);

            foreach (var order in orders)
            {
                order.ArchivPath =  "Q:\\Archiv\\Technical_Functions\\420_Musterbau\\10J\\" + order.Aid;
                db.Update(order);

            }
            db.SaveChanges();
            //var pcont = new PersonalFilterContainer();
            //var filt = new PersonalFilter("^F", PropertyNames.Auftragsnummer);
            //pcont.Add("name", filt);
            //var res = filt.TestValue(new Vorgang() { Aid = "f2100", BemM = "V" });
            //using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            //var vorg = db.Vorgangs;

            //foreach (var v in vorg.Skip(40).Take(10))
            //{
            //    var res = new Response() { Rework = 0, Scrap = 0, Yield = 112, Timestamp = DateTime.Now };
            //    v.Responses.Add(res);
            //}
            //db.SaveChanges();

            //string string1 = "<row><VorgangID>0060082254</VorgangID><AID>G2024_0182</AID><VNR>80</VNR><ArbPlSAP>20120100</ArbPlSAP><Text>Zählen, Verpacken, Abliefern</Text><SpaetStart>2025-01-30T00:00:00</SpaetStart><SpaetEnd>2025-01-30T00:00:00</SpaetEnd><BEAZE>0.0000000e+000</BEAZE><SysStatus>FREI TRÜC RMST RMAN </SysStatus><SteuSchl>PPKB</SteuSchl><BasisMg>1</BasisMg><aktuell>1</aktuell><Quantity-scrap>0</Quantity-scrap><Quantity-yield>69</Quantity-yield><Quantity-miss>14</Quantity-miss><ProcessingUOM>MIN             </ProcessingUOM><Quantity-rework>0</Quantity-rework><Visability>1</Visability><ActualStartDate>2024-11-13T00:00:00</ActualStartDate><ActualEndDate>2024-11-20T00:00:00</ActualEndDate><Bullet>#FFFFFFFF</Bullet><Quantity-miss-neo>280</Quantity-miss-neo></row>";
            //string string2 = "<row><VorgangID>0060082254</VorgangID><AID>G2024_0182</AID><VNR>80</VNR><ArbPlSAP>20120100</ArbPlSAP><Text>Zählen, Verpacken, Abliefern</Text><SpaetStart>2025-01-30T00:00:00</SpaetStart><SpaetEnd>2025-01-30T00:00:00</SpaetEnd><BEAZE>0.0000000e+000</BEAZE><SysStatus>FREI TRÜC RMST RMAN </SysStatus><SteuSchl>PPKB</SteuSchl><BasisMg>1</BasisMg><aktuell>0</aktuell><Quantity-scrap>0</Quantity-scrap><Quantity-yield>83</Quantity-yield><Quantity-miss>0</Quantity-miss><ProcessingUOM>MIN             </ProcessingUOM><Quantity-rework>0</Quantity-rework><Visability>1</Visability><ActualStartDate>2024-11-13T00:00:00</ActualStartDate><ActualEndDate>2024-11-22T00:00:00</ActualEndDate><Bullet>#FFFFFFFF</Bullet><Quantity-miss-neo>266</Quantity-miss-neo></row>";

            //var xml1 = XElement.Parse(string1);
            //var xml2 = XElement.Parse(string2);
            //StringBuilder result = new StringBuilder();
            //foreach ( var x in xml1.Elements())
            //{
            //    if ( x.Value != xml2.Element(x.Name).Value )
            //    {
            //        result.Append(x.Name).Append(": ").Append(x.Value).Append(" => ").Append(xml2.Element(x.Name).Value);
            //    }
            //}

        }

    }
}


