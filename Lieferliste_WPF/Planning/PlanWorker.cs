using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using GongSolutions.Wpf.DragDrop;
using Lieferliste_WPF.Utilities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace Lieferliste_WPF.Planning
{
    public interface IPlanWorkerFactory
    {
        IContainerProvider Container { get; }
        IApplicationCommands ApplicationCommands { get; }
        IUserSettingsService SettingsService { get; }
    }
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class PlanWorkerFactory : IPlanWorkerFactory
    {
        public IContainerProvider Container { get; }

        public IApplicationCommands ApplicationCommands { get; }

        public IUserSettingsService SettingsService { get; }
        public IDialogService DialogService { get; }

        public PlanWorkerFactory(IContainerProvider container, IApplicationCommands applicationCommands,
            IUserSettingsService settingsService, IDialogService dialogService)
        {
            this.Container = container;
            this.ApplicationCommands = applicationCommands;
            this.SettingsService = settingsService;
            this.DialogService = dialogService;
            
        }
        public PlanWorker CreatePlanWorker(int MessRid, List<Vorgang> processesAll)
        {
            return new PlanWorker(MessRid, processesAll, Container, ApplicationCommands, SettingsService, DialogService);
        }
    }
    public interface IPlanWorker
    {
        public int MessRId { get; }

    }
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class PlanWorker : DependencyObject, IPlanWorker, IDropTarget, IViewModel
    {

        #region Constructors

        public PlanWorker(int MessRId, List<Vorgang> processes, IContainerProvider container,
            IApplicationCommands applicationCommands,
            IUserSettingsService settingsService,
            IDialogService dialogService)
        {
            _container = container;
            _messRid = MessRId;
            _applicationCommands = applicationCommands;
            _settingsService = settingsService;
            _dialogService = dialogService;
            EnableRowDetails = settingsService.IsRowDetails;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<PlanWorker>();
            Initialize();
            Processes.AddRange(processes.Where(x => x.MsfInWork == false));
            LoadData();

        }

        #endregion

        public ICommand? SetMarkerCommand { get; private set; }
        public ICommand? OpenWorkerCommand { get; private set; }
        public ICommand? WorkerPrintCommand { get; private set; }
        public ICommand? KlimaPrintCommand { get; private set; }
        public ICommand? DocumentAddCommand { get; private set; }
        private readonly int _messRid;
        ILogger logger;
        public int MessRId => _messRid;
        public string? Name { get; private set; }
        public string UserId { get; private set; }
        public bool EnableRowDetails { get; }
        public string? Description { get; private set; }
        public WorkArea? WorkArea { get; set; }
        public List<int> CostUnits { get; set; } = [];
        public ObservableCollection<Vorgang>? Processes { get; set; }

        public ICollectionView ProcessesCV { get { return ProcessesCVSource.View; } }
        private IContainerProvider _container;
        private IUserSettingsService _settingsService;
        private IDialogService _dialogService;
        private IApplicationCommands? _applicationCommands;

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
        private string _title = "Messtechnik";
        public string Title => _title;

        public bool HasChange => throw new NotImplementedException();

        private void LoadData()
        {
            try
            {
                using var _dbctx = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var usr = _dbctx.MeasureResses.Include(x => x.User).First(x => x.MessRid == MessRId).User;

                Name = string.Format("{0} {1}", usr.Firstname, usr.Lastname);
                Description = usr.Department;
                UserId = usr.AccountId;
                ProcessesCV.SortDescriptions.Add(new SortDescription("SortPos", ListSortDirection.Ascending));
                ProcessesCV.Filter = f => !((Vorgang)f).SysStatus?.Contains("RÜCK") ?? false;
                ProcessesCV.Refresh();
                var live = ProcessesCV as ICollectionViewLiveShaping;
                if (live != null)
                {
                    live.IsLiveSorting = false;
                    live.LiveFilteringProperties.Add("SysStatus");
                    live.IsLiveFiltering = true;
                }
            }
            catch (Exception e)
            {
                logger.LogError("{message}", e.ToString());
            }
        }
        private void Initialize()
        {
            SetMarkerCommand = new ActionCommand(OnSetMarkerExecuted, OnSetMarkerCanExecute);
            WorkerPrintCommand = new ActionCommand(OnWorkerPrintExecuted, OnWorkerPrintCanExecute);
            OpenWorkerCommand = new ActionCommand(OnOpenWorkerExecuted, OnOpenWorkerCanExecute);
            KlimaPrintCommand = new ActionCommand(OnKlimaPrintExecuted, OnKlimaPrintCanExecute);
            DocumentAddCommand = new ActionCommand(OnDocumentAddExecuted, OnDocumentAddCanExecute);
            Processes = new ObservableCollection<Vorgang>();
            ProcessesCVSource.Source = Processes;
        }

        private bool OnOpenWorkerCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.MessWorker);
        }

        private void OnOpenWorkerExecuted(object obj)
        {
            try
            {
                var par = new DialogParameters();
                par.Add("PlanWorker", this);

                _dialogService.Show("WorkerView", par, WorkerViewCallBack);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error OpenWorker", MessageBoxButton.OK, MessageBoxImage.Error);
                logger.LogError("{message}", e.ToString());
            }
        }

        private void WorkerViewCallBack(IDialogResult result)
        {

        }

        private bool OnDocumentAddCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddMeasureDocu);
        }

        private void OnDocumentAddExecuted(object obj)
        {
            if (obj is Vorgang vrg)
            {
                _dialogService.Show("DocumentDialog", DocumentCallBack);
            }
        }

        private void DocumentCallBack(IDialogResult result)
        {

        }

        private bool OnKlimaPrintCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.KlimaPrint);
        }

        private void OnKlimaPrintExecuted(object obj)
        {
            try
            {
                if (obj is Vorgang vrg)
                {
                    bool print = true;
                    if (!vrg.KlimaPrint.HasValue)
                    {
                        vrg.KlimaPrint = DateTime.Now;

                    }
                    else
                    {
                        var result = MessageBox.Show(string.Format("Es wurde bereits am {0} ausgedruckt.\nSoll nochmals gedruckt werden?",
                            vrg.KlimaPrint?.ToString("dd/MM/yy HH:mm")),
                            "Info Klimaausduck", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (result == MessageBoxResult.No) { print = false; }
                    }
                    if (print)
                    {
                        var fd = Printing.CreateKlimaDocument(vrg);
                        PrintTicket ticket = new PrintTicket();
                        ticket.PageOrientation = PageOrientation.Landscape;
                        ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA5);

                        Printing.DoThePrint(fd, ticket, vrg.VorgangId + "-" + vrg.KlimaPrint?.ToString("ddMMyyHHmm"));
                        logger.LogInformation("{message}", vrg.VorgangId);
                    }
                }
            }
            catch (Exception e)
            {
                logger.LogError("{message}", e.ToString());
            }
        }

        private bool OnWorkerPrintCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.MachPrint);
        }

        private void OnWorkerPrintExecuted(object obj)
        {
            try
            {
                var print = new PrintDialog();
                var ticket = new PrintTicket();
                ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
                ticket.PageOrientation = PageOrientation.Portrait;
                print.PrintTicket = ticket;
                //Printing.DoPrintPreview(obj, print);
                logger.LogInformation("{message}", obj.ToString());
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message, "WorkerPrint", MessageBoxButton.OK, MessageBoxImage.Error);
                logger.LogError("{message}", e.ToString());
            }
        }

        private static bool OnSetMarkerCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.MessSetMark);
        }

        private void OnSetMarkerExecuted(object obj)
        {
            try
            {
                if (obj == null) return;
                var values = (object[])obj;
                var name = (string)values[0];
                var desc = (Vorgang)values[1];

                if (name == "DelBullet") desc.Bullet = Brushes.White.ToString();
                if (name == "Bullet1") desc.Bullet = Brushes.Red.ToString();
                if (name == "Bullet2") desc.Bullet = Brushes.Green.ToString();
                if (name == "Bullet3") desc.Bullet = Brushes.Yellow.ToString();
                if (name == "Bullet4") desc.Bullet = Brushes.Blue.ToString();

                ProcessesCV.Refresh();
                logger.LogInformation("{message}", desc.VorgangId);
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message, "SetMarkerMess", MessageBoxButton.OK, MessageBoxImage.Error);
                logger.LogError("{message}", e.ToString());
            }

        }

        public void Drop(IDropInfo dropInfo)
        {

            try
            {
                var vrg = (Vorgang)dropInfo.Data;
                var s = dropInfo.DragInfo.SourceCollection as ListCollectionView;
                var t = dropInfo.TargetCollection as ListCollectionView;
                if (s.CanRemove) s.Remove(vrg);
                var v = dropInfo.InsertIndex;
                if(vrg.MsfInWork) vrg.MsfInWork = false;
                if (v > t?.Count)
                {
                    ((IList)t.SourceCollection).Add(vrg);
                }
                else
                {
                    ((IList)t.SourceCollection).Insert(v, vrg);
                }
                var p = t.SourceCollection as Collection<Vorgang>;

                for (var i = 0; i < p.Count; i++)
                {
                    p[i].SortPos = (p[i].SysStatus?.Contains("RÜCK") == true) ? "Z" :
                        string.Format("{0}_{1,3:0}", MessRId, i.ToString("D3"));
                }
                var msid = MessRId.ToString().Concat(vrg.VorgangId);
                if (vrg.MeasureRessVorgangs.All(x => x.MessId.ToString().Concat(x.VorgId) != msid))
                {
                    var existing = vrg.MeasureRessVorgangs.FirstOrDefault(x => x.VorgId == vrg.VorgangId);
                    if (existing != null) { vrg.MeasureRessVorgangs.Remove(existing); }
                    vrg.MeasureRessVorgangs.Add(new MeasureRessVorgang() { MessId = this.MessRId, VorgId = vrg.VorgangId });
                }
                logger.LogInformation("drops {message}", vrg.VorgangId);
                t.Refresh();

            }
            catch (Exception e)
            {
                string str = string.Format("{0}\n{1}", e.Message, e.InnerException);
                MessageBox.Show(str, "ERROR", MessageBoxButton.OK);
                logger.LogError("{message}", e.ToString());
            }
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

        public void Closing()
        {
        }
    }
}
