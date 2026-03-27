using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.Extensions.Logging;
using ModuleShift.Services;
using Prism.Ioc;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml;
using static El2Core.Constants.ShiftTypes;

namespace ModuleShift.ViewModels
{
    internal class ShiftEditViewModel : ViewModelBase
    {
        public ShiftEditViewModel(IContainerExtension container)
        {
            _container = container;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<ShiftEditViewModel>();
            SaveCommand = new ActionCommand(onSaveExecuted, onSaveCanExecute);
            AddCommand = new ActionCommand(onAddExecuted, onAddCanExecute);
            RemoveCommand = new ActionCommand(onRemoveExecuted, onRemoveCanExecute);
           
            LoadData();
            ShiftView = CollectionViewSource.GetDefaultView(WorkShiftCollection);
            ShiftView.MoveCurrentToFirst();
        }

        public string Title { get; } = "Schicht Editor";
        IContainerExtension _container;
        ILogger logger;
        public ObservableCollection<WorkShiftService> WorkShiftCollection { get; private set; } = [];
        public ObservableCollection<DataTable> ShiftDataSets { get; private set; } = [];
        private ObservableCollection<WorkShiftItem> WorkShiftItems = [];
        private bool isAdding;
        public bool IsAdding
        {
            get
            {
                return isAdding;
            }
            set
            {
                if (value != isAdding)
                {
                    isAdding = value;
                    NotifyPropertyChanged(() => IsAdding);
                }
            }
        }

        public DataSet dataSet { get; private set; }
        public DataView dataView { get; private set; }
        public ICollectionView ShiftView { get; private set; }
        public ICollectionView ShiftItemsView { get; private set; }
        public string xmlString { get; set; }
        public XmlDataProvider xmlDataProvider { get; private set; }
        public ICommand SaveCommand { get; private set; }
        public ICommand AddCommand { get; private set; }
        public ICommand RemoveCommand { get; private set; }
        private RelayCommand? lostFocusCommand;
        public RelayCommand LostFocusCommand => lostFocusCommand ??= new RelayCommand(onLostFocus);


        private void LoadData()
        {
            try
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var ws = db.WorkShifts.ToList();
                var serializer = XmlSerializerHelper.GetSerializer(typeof(ObservableCollection<WorkShiftItem>));
                foreach (var w in ws)
                {
                    using XmlReader reader = XmlReader.Create(new StringReader(w.ShiftDef));
                    var wos = new WorkShiftService()
                    {
                        ShiftName = w.ShiftName,
                        id = w.Sid,
                        ShiftType = (ShiftType)w.ShiftType,
                        Items = (ObservableCollection<WorkShiftItem>)serializer.Deserialize(reader)
                    };
                    wos.Changed = false;
                    foreach (var item in wos.Items)
                    {
                        item.PropertyChanged += onShiftItemChanged;
                    }
                    wos.PropertyChanged += onWorkShiftServiceChanged;
                    wos.Items.CollectionChanged += onItemsCollectionChanged;
                    WorkShiftCollection.Add(wos);
                }
            }
            catch (System.Exception e)
            {

                logger.LogError("{message}", e.ToString());
            }
        }

        private void onItemsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            var items = sender as ObservableCollection<WorkShiftItem>;
            if (items != null)
                items.Last().Changed = true;
        }

        private void onShiftItemChanged(object? sender, PropertyChangedEventArgs e)
        {
            var item = sender as WorkShiftItem;
            if (item != null)
                item.Changed = true;
        }

        private void onWorkShiftServiceChanged(object? sender, PropertyChangedEventArgs e)
        {
            var ws = sender as WorkShiftService;
            if (ws != null)
                ws.Changed = true;
        }

        private bool onSaveCanExecute(object arg)
        {
            var serv = WorkShiftCollection.Any(x => x.Changed);
            var item = WorkShiftCollection.Where(x => x.Items.Any(y => y.Changed)).Any();
            return serv || item;
        }

        private void onSaveExecuted(object obj)
        {
            try
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var wo = db.WorkShifts;
                var serializer = XmlSerializerHelper.GetSerializer(typeof(ObservableCollection<WorkShiftItem>));

                foreach (var w in WorkShiftCollection)
                {
                    var sw = new StringWriter();
                    if (w.id == 0)
                    {
                        serializer.Serialize(sw, w.Items);
                        var work = new WorkShift() { ShiftName = w.ShiftName, ShiftType = (int)w.ShiftType, ShiftDef = sw.ToString() };

                        wo.Add(work);
                    }
                    else if (w.Changed)
                    {
                        var sh = wo.Find(w.id);
                        sh.ShiftName = w.ShiftName;
                        sh.ShiftType = (int)w.ShiftType;
                        w.Changed = false;
                    }
                    if (w.Items.Any(x => x.Changed))
                    {
                        serializer.Serialize(sw, w.Items);
                        var work = wo.First(x => x.Sid == w.id);
                        work.ShiftName = w.ShiftName;
                        work.ShiftType = (int)w.ShiftType;
                        work.ShiftDef = sw.ToString();
                    }
                    foreach (var item in w.Items.Where(x => x.Changed))
                    {
                        item.Changed = false;
                    }
                }
                db.SaveChanges();
                logger.LogInformation("{message}", [.. wo]);
            }
            catch (System.Exception e)
            {

                logger.LogError("{message}", e.ToString());
            }

        }

        private bool onRemoveCanExecute(object arg)
        {
            return true;
        }

        private void onRemoveExecuted(object obj)
        {
            if(obj is WorkShiftItem wsi)
            {
                ((WorkShiftService)ShiftView.CurrentItem).Items.Remove(wsi);
            }
            if(obj is WorkShiftService ws)
            {
                WorkShiftCollection.Remove(ws);
            }
        }

        private bool onAddCanExecute(object arg)
        {
            bool accept = false;
            var current = (WorkShiftService)ShiftView.CurrentItem;
            if (arg == null)
            {
                accept = string.IsNullOrEmpty(current.ShiftName) == false && current.Items.All(x => x.EndTime != null && x.StartTime != null);
            }
            else if (arg is string and "ITEM")
                accept = current.Items.All(x => x.EndTime != null && x.StartTime != null);
            return accept;
        }

        private void onAddExecuted(object obj)
        {
            IsAdding = true;
            if (obj == null)
            {
                WorkShiftCollection.Add(new WorkShiftService());
                ShiftView.MoveCurrentToLast();
            }
            WorkShiftService e = (WorkShiftService)ShiftView.CurrentItem;
            e.Items.Add(new WorkShiftItem());
        }
        private void onLostFocus(object obj)
        {
            var e = ShiftView.CurrentItem;
            var v = WorkShiftCollection.Count;
        }
    }
}
