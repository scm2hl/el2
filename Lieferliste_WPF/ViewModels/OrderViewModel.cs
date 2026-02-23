
using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Prism.Dialogs;
using Prism.Events;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace Lieferliste_WPF.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class OrderViewModel : ViewModelBase, IDialogAware
    {
        public OrderViewModel(IContainerProvider container, IApplicationCommands applicationCommands, IEventAggregator ea)
        {
            _applicationCommands = applicationCommands;
            _ea = ea;
            _container = container;
            VorgangCV = CollectionViewSource.GetDefaultView(Vorgangs);
            SaveCommand = new ActionCommand(OnSaveExecuted, OnSaveCanExecute);
            DeleteCommand = new ActionCommand(OnDeleteExecuted, OnDeleteCanExecute);
            _ea.GetEvent<MessageVorgangChanged>().Subscribe(OnMessageReceived);
            CopyClipBoardCommand = new ActionCommand(OnCopyClipBoardExecuted, OnCopyClipBoardCanExecute);
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

        }

        private bool OnCopyClipBoardCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.CopyClipboard);
        }

        private void OnCopyClipBoardExecuted(object obj)
        {
            if (obj is OrderViewModel o)
            {
                Clipboard.SetText(string.Format("Auftragsnummer: {0}\nMaterial: {1} {2}\nMenge: {3}\nTermin: {4:D}",
                    o.Aid, o.Material, o.Bezeichnung, o.Quantity, o.EckEnde));
            }
        }

        private void OnMessageReceived(List<(string, string)?> vorgangIdList)
        {
            try
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                foreach (var vid in vorgangIdList)
                {
                    var vo = Vorgangs.FirstOrDefault(x => x.VorgangId == vid.Value.Item2);

                    if (vo != null)
                    {
                        var v = db.Vorgangs.First(x => x.VorgangId == vo.VorgangId);

                        if (vid.Value.Item1 == "EL2")
                        {
                            vo.BemM = v.BemM;
                            vo.BemMa = v.BemMa;
                            vo.BemT = v.BemT;
                        }
                        else
                        {
                            vo.SysStatus = v.SysStatus;
                            vo.QuantityMiss = v.QuantityMiss;
                            vo.QuantityRework = v.QuantityRework;
                            vo.QuantityScrap = v.QuantityScrap;
                            vo.QuantityYield = v.QuantityYield;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "MsgReceivedOrderView", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public ICommand CopyClipBoardCommand { get; }
        private IContainerProvider _container;
        private readonly IEventAggregator _ea;
        private bool OnSaveCanExecute(object arg)
        {
            return false;
        }

        private void OnSaveExecuted(object obj)
        {

        }
        private bool OnDeleteCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.LieferVrgDel);
        }

        private void OnDeleteExecuted(object obj)
        {
            if (obj is Vorgang v)
            {
                var result = MessageBox.Show(string.Format("Soll der Vorgang {0:d4}\n vom Auftrag {1} gelöscht werden?", v.Vnr, v.Aid), "Information",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    db.Vorgangs.Remove(v);
                    db.SaveChanges();
                }
            }
        }

        private ObservableCollection<Vorgang> Vorgangs { get; } = [];
        private RelayCommand? _BemChangedCommand;
        public RelayCommand BemChangedCommand => _BemChangedCommand ??= new RelayCommand(OnBemChanged);

        private void OnBemChanged(object obj)
        {
            if (obj is Vorgang vrg)
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var v = db.Vorgangs.Single<Vorgang>(x => x.VorgangId == vrg.VorgangId);
                if (v.BemM != vrg.BemM)
                {
                    _ = Globals.NotifyBroker.SendMessageAsync(vrg.BemM, "MaBem");
                    v.BemM = vrg.BemM;
                }
                v.BemMa = vrg.BemMa;
                v.BemT = vrg.BemT;

                db.SaveChanges();
            }
        }
        private RelayCommand? _DateChangedCommand;
        public RelayCommand DateChangedCommand => _DateChangedCommand ??= new RelayCommand(OnDateChanged);
        public ActionCommand DeleteCommand { get; private set; }
        private void OnDateChanged(object obj)
        {
            if (obj is Vorgang vrg)
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var v = db.Vorgangs.Single<Vorgang>(x => x.VorgangId == vrg.VorgangId);
                v.Termin = vrg.Termin;

                db.SaveChanges();
            }
        }
        public string UserIdent { get; } = UserInfo.User.UserId;
        private string _aid;
        public string Aid
        {
            get { return _aid; }
            set
            {
                if (_aid != value)
                {
                    _aid = value;
                    NotifyPropertyChanged(() => Aid);
                }
            }
        }
        private DateTime? _eckEnde;

        public DateTime? EckEnde
        {
            get { return _eckEnde; }
            set
            {
                if (_eckEnde != value)
                {
                    _eckEnde = value;
                    NotifyPropertyChanged(() => EckEnde);
                }
            }
        }

        private string? _material;
        public string? Material
        {
            get { return _material; }
            set
            {
                if (_material != value)
                {
                    _material = value;
                    NotifyPropertyChanged(() => Material);
                }
            }
        }
        private string? _dummyMat;
        public string? DummyMat
        {
            get { return _dummyMat; }
            set
            {
                if (_dummyMat != value)
                {
                    _dummyMat = value;
                    NotifyPropertyChanged(() => DummyMat);
                }
            }
        }
        private string? _bezeichnung;
        public string? Bezeichnung
        {
            get { return _bezeichnung; }
            set
            {
                if (_bezeichnung != value)
                {
                    _bezeichnung = value;
                    NotifyPropertyChanged(() => Bezeichnung);
                }
            }
        }
        private int? _quantity;

        public int? Quantity
        {
            get { return _quantity; }
            set
            {
                if (_quantity != value)
                {
                    _quantity = value;
                    NotifyPropertyChanged(() => Quantity);
                }
            }
        }

        private string? _pro;

        public string? Pro
        {
            get { return _pro; }
            set
            {
                if (value != _pro)
                {
                    _pro = value;
                    NotifyPropertyChanged(() => Pro);
                }
            }
        }
        private string? _proInfo;

        public string? ProInfo
        {
            get { return _proInfo; }
            set
            {
                if ((_proInfo != value))
                {
                    _proInfo = value;
                    NotifyPropertyChanged(() => ProInfo);
                }
            }
        }
        private string[] _msflist;
        public string[] MSFList
        {
            get { return _msflist; }
            private set
            {
                _msflist = value;
                NotifyPropertyChanged(() => MSFList);
            }
        }

        private string? _sysStatus;

        public string? SysStatus
        {
            get { return _sysStatus; }
            set
            {
                if (_sysStatus != value)
                {
                    _sysStatus = value;
                    NotifyPropertyChanged(() => SysStatus);
                }
            }
        }
        private bool _ready;

        public bool Ready
        {
            get { return _ready; }
            set
            {
                if (_ready != value)
                {
                    _ready = value;
                    NotifyPropertyChanged(() => Ready);
                }
            }
        }
        private int _archivState = 0;

        public int ArchivState
        {
            get { return _archivState; }
        }
        private string? _ArchivPath;
        public string? ArchivPath { get { return _ArchivPath; } }

        public ICollectionView VorgangCV { get; }
        private string _title = "Auftragsübersicht";
        public string Title
        {
            get { return _title; }
            set
            {
                if (_title != value)
                {
                    _title = value;
                    NotifyPropertyChanged(() => Title);
                }
            }
        }
        private IApplicationCommands _applicationCommands;
        public IApplicationCommands ApplicationCommands
        {
            get { return _applicationCommands; }
            set
            {
                if (_applicationCommands != value)
                    _applicationCommands = value;
                NotifyPropertyChanged(() => ApplicationCommands);
            }
        }

        public DialogCloseListener RequestClose { get; }

        public ICommand SaveCommand;

        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {

        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            var p = parameters.GetValue<List<Vorgang>>("vrgList");

            if (p.Count > 0)
            {
                foreach (var item in p.OrderBy(x => x.Vnr))
                {
                    Vorgangs.Add(item);
                }
                var v = p.First();
                Aid = v.Aid;
                Material = v.AidNavigation.Material ?? "DUMMY";
                if (Material == "DUMMY") DummyMat = v.AidNavigation.DummyMat;
                Bezeichnung = (string.IsNullOrEmpty(v.AidNavigation.Material)) ? v.AidNavigation.DummyMatNavigation?.Mattext : v.AidNavigation.MaterialNavigation?.Bezeichng;
                Quantity = v.AidNavigation.Quantity;
                Pro = v.AidNavigation.ProId;
                ProInfo = v.AidNavigation.Pro?.ProjectInfo;
                SysStatus = v.AidNavigation.SysStatus;
                Ready = v.AidNavigation.Fertig;
                EckEnde = v.AidNavigation.Eckende;
                _archivState = v.AidNavigation.ArchivState;
                _ArchivPath = v.AidNavigation.ArchivPath;
                VorgangCV.Refresh();
            }
            else
                MessageBox.Show("keine Vorgänge vorhanden", Title, MessageBoxButton.OK, MessageBoxImage.Error);

            MSFList = p.Where(x => string.IsNullOrEmpty(x.Msf) == false).Select(y => y.Msf).ToArray();
        }
    }
}
