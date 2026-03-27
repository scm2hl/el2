using El2Core.Constants;
using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using GongSolutions.Wpf.DragDrop;
using Microsoft.EntityFrameworkCore;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;


namespace ModuleShift.ViewModels
{
    internal class ShiftPlanEditViewModel : ViewModelBase, IDropTarget
    {
        private IContainerProvider _container;
        private IDialogService _dialogService;
        public string Title { get; } = "Schichtplan";
        private RelayCommand? _ShiftPlanSelectionChangedCommand;
        public RelayCommand ShiftPlanSelectionChangedCommand => _ShiftPlanSelectionChangedCommand ??= new RelayCommand(OnPlanSelected);
        private RelayCommand? _ComboSelectionChangedCommand;
        public RelayCommand ComboSelectionChangedCommand => _ComboSelectionChangedCommand ??= new RelayCommand(OnSelectionChanged);

        public ICommand SaveAllCommand { get; private set; }
        public ICommand DeleteCommand { get; private set;}
        public ICommand SaveNewCommand { get; private set; }
        public ICommand DetailCoverCommand { get; private set; }
        public ICommand AddCoverCommand { get; private set; }
        public ICommand NewCalendarCommand { get; private set; }
        public ICommand DelCalendarCommand { get; private set; }
        public ICommand NewCalendarShiftCommand { get; private set; }
        public ICommand DelCalendarShiftCommand { get; private set; }
        public Dictionary<int, List<ShiftDay>> ShiftWeeks { get; set; }
        private ShiftWeek _SelectedPlan;
        public ShiftWeek SelectedPlan
        {
            get { return _SelectedPlan; }
            set
            {
                if (_SelectedPlan != value)
                {
                    _SelectedPlan = value;
                    NotifyPropertyChanged(() => SelectedPlan);
                }
            }
        }
        private ShiftCalendar selectedCalendar;

        public ShiftCalendar SelectedCalendar
        {
            get { return selectedCalendar; }
            set
            {
                if (selectedCalendar != value)
                {
                    selectedCalendar = value;
                    NotifyPropertyChanged(() => SelectedCalendar);
                }
            }
        }
        private List<ShiftWeek> _ShiftWeekPlans { get; set; }
        public ICollectionView ShiftWeekPlans { get; private set; }
        private ObservableCollection<ShiftCalendar> shiftCalendars = [];
        public ICollectionView ShiftCalendars { get; private set; }
        public ShiftPlanEditViewModel(IContainerProvider container, IDialogService dialogService)
        {
            _container = container;
            _dialogService = dialogService;
            SaveAllCommand = new ActionCommand(OnSaveAllExecuted, OnSaveAllCanExecute);
            SaveNewCommand = new ActionCommand(OnSaveNewExecuted, OnSaveNewCanExecute);
            DeleteCommand = new ActionCommand(OnDeleteExecuted, OnDeleteCanExecuted);
            AddCoverCommand = new ActionCommand(OnAddExecuted, OnAddCanExecuted);
            DetailCoverCommand = new ActionCommand(OnDetailExecuted, OnDetailCanExecuted);
            NewCalendarCommand = new ActionCommand(OnNewCalendarExecuted, OnNewCalendarCanExecute);
            DelCalendarCommand = new ActionCommand(OnDelCalendarExecuted, OnDelCalendarCanExecute);
            NewCalendarShiftCommand = new ActionCommand(OnNewCalendarShiftExecuted, OnNewCalendarShiftCanExecute);
            DelCalendarShiftCommand = new ActionCommand(OnDelCalendarShiftExecuted, OnDelCalendarShiftCanExecute);
            LoadData();
            LoadCovers();
            
        }


        public bool IsRubberChecked { get; set; }
        private List<ShiftCover> _ShiftCovers = [];
        public ICollectionView ShiftCovers { get; private set; }
        private void LoadCovers()
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var shc = db.ShiftCovers;
 
            _ShiftCovers.AddRange(shc);
            ShiftCovers = CollectionViewSource.GetDefaultView(_ShiftCovers);

        }
        private void SaveShiftCover(ShiftCover shiftCover)
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var shc = db.ShiftCovers;

            if(shc.Contains(shiftCover))
            {

            }
            else
            {
                shc.Add(shiftCover);
            }
            db.SaveChanges();
        }
        private void LoadData()
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var shift = db.ShiftPlans.Where(x => x.Id != 0).AsNoTracking().OrderBy(x => x.Id);
            if (shift.Any())
            {

                ShiftWeeks = [];
                _ShiftWeekPlans = [];
                foreach (var item in shift)
                {
                    var week = new ShiftWeek();
                    week.Id = item.Id;
                    week.ShiftPlanName = item.PlanName;
                    week.Lock = item.Lock;
                   
                    List<ShiftDay> shiftDays = new();
                    Byte[] bytes;
                    bytes = item.Sun;
                    week.ShiftWeekDays.Add(new ShiftDay(0, new BitArray(bytes)));
                    bytes = item.Mon;
                    week.ShiftWeekDays.Add(new ShiftDay(1, new BitArray(bytes)));
                    bytes = item.Tue;
                    week.ShiftWeekDays.Add(new ShiftDay(2, new BitArray(bytes)));
                    bytes = item.Wed;
                    week.ShiftWeekDays.Add(new ShiftDay(3, new BitArray(bytes)));
                    bytes = item.Thu;
                    week.ShiftWeekDays.Add(new ShiftDay(4, new BitArray(bytes)));
                    bytes = item.Fre;
                    week.ShiftWeekDays.Add(new ShiftDay(5, new BitArray(bytes)));
                    bytes = item.Sat;
                    week.ShiftWeekDays.Add(new ShiftDay(6, new BitArray(bytes)));

                    _ShiftWeekPlans.Add(week);
                }
                SelectedPlan = _ShiftWeekPlans.First();
                ShiftWeekPlans = CollectionViewSource.GetDefaultView(_ShiftWeekPlans);
                var cal = db.ShiftCalendars
                    .Include(x => x.ShiftCalendarShiftPlans)
                    .ToList();
                foreach (var scal in cal)
                {
                    var c = new ShiftCalendar() { id = scal.Id, CalendarName = scal.CalendarName, IsLocked = scal.Lock, Repeat = scal.Repeat };
                    foreach (var sc in scal.ShiftCalendarShiftPlans.OrderBy(x => x.YearKw))
                    {
                        var sw = (ShiftWeek)_ShiftWeekPlans.Single(x => x.Id == sc.PlanId).Clone();
                        sw.YearKW = sc.YearKw;
                        c.ShiftWeeks.Add(sw);
                    }
                    c.ShiftWeeks.CollectionChanged += CalendarShiftWeeksChanged;
                    shiftCalendars.Add(c);
                }
                ShiftCalendars = CollectionViewSource.GetDefaultView(shiftCalendars);
                SelectedCalendar = (ShiftCalendar)ShiftCalendars.CurrentItem;
            }
        }

        private void CalendarShiftWeeksChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            SelectedCalendar.IsChanged = true;
        }

        private Byte[]? GetDefinition(ShiftPlan shiftPlanDb, int index)
        {
            switch (index)
            {
                case 0: return shiftPlanDb.Sun;
                case 1: return shiftPlanDb.Mon;
                case 2: return shiftPlanDb.Tue;
                case 3: return shiftPlanDb.Wed;
                case 4: return shiftPlanDb.Thu;
                case 5: return shiftPlanDb.Fre;
                case 6: return shiftPlanDb.Sat;
            }
            return null;
        }

        private bool OnDelCalendarShiftCanExecute(object arg)
        {
            return !SelectedCalendar.IsLocked &&
                PermissionsProvider.GetInstance().GetUserPermission(Permissions.DelShiftCalendar);
        }

        private void OnDelCalendarShiftExecuted(object obj)
        {
            if (obj is ShiftWeek sw)
            {
                SelectedCalendar.ShiftWeeks.Remove(sw);
                SelectedCalendar.IsChanged = true;
            }
        }

        private bool OnNewCalendarShiftCanExecute(object arg)
        {
            return !SelectedCalendar.IsLocked &&
                PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddShiftCalendar);
        }

        private void OnNewCalendarShiftExecuted(object obj)
        {
            var sw = (ShiftWeek)_ShiftWeekPlans.First().Clone();
            var s = SelectedCalendar.ShiftWeeks.LastOrDefault();
            int wo = 0;
            DateTime sdt = DateTime.Now;
            if (s != null) { sdt = new DateTime(int.Parse(s.YearKW[..4]), 1, 1); wo = int.Parse(s.YearKW[4..]); }
            sdt = CultureInfo.InvariantCulture.Calendar.AddWeeks(sdt, wo);
            sw.YearKW = string.Format("{0}{1}", sdt.Year,
                       CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(sdt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday));
            
            SelectedCalendar.ShiftWeeks.Add(sw);
            SelectedCalendar.IsChanged = true;
        }

        private bool OnDelCalendarCanExecute(object arg)
        {
            return !SelectedCalendar.IsLocked &&
                PermissionsProvider.GetInstance().GetUserPermission(Permissions.DelShiftCalendar);
        }

        private void OnDelCalendarExecuted(object obj)
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var c = db.ShiftCalendars.Single(x => x.Id == SelectedCalendar.id);
            db.ShiftCalendars.Remove(c);
            db.SaveChangesAsync();
            shiftCalendars.Remove(SelectedCalendar);
            SelectedCalendar = shiftCalendars.First();           
        }

        private bool OnNewCalendarCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddShiftCalendar);
        }

        private void OnNewCalendarExecuted(object obj)
        {
            _dialogService.ShowDialog("InputDialog", DialogCallBack);
        }

        private void DialogCallBack(IDialogResult result)
        {
            if (result.Result == ButtonResult.OK)
            {
                var r = result.Parameters.GetValue<string>("InputText");
                
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var sc = new El2Core.Models.ShiftCalendar() { CalendarName = r };
                db.ShiftCalendars.Add(sc);
                db.SaveChanges();
                SelectedCalendar = new ShiftCalendar() { CalendarName = r, id = sc.Id };
                shiftCalendars.Add(SelectedCalendar);
            }
        }

        private bool OnDetailCanExecuted(object arg)
        {
            return true;
        }

        private void OnDetailExecuted(object obj)
        {
            if (obj is ShiftCover cover)
            {
                var par = new DialogParameters();
                par.Add("Cover", cover);

                _dialogService.Show("DetailCoverDialog", par, OnDetailCallBack);
            }
        }

        private void OnDetailCallBack(IDialogResult result)
        {
            if (result.Result == ButtonResult.OK)
            {
                var c = result.Parameters.GetValue<ShiftCover>("Cover");
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                if (string.IsNullOrEmpty(c.CoverName) == false &&
                        c.CoverMask != null)
                {
                    
                    if (c.Id == 0 )
                    {
                        db.ShiftCovers.Add(c);
                        _ShiftCovers.Add(c);
                        ShiftCovers.Refresh();
                    }
                    else
                    {
                        db.ShiftCovers.Update(c);
                        var cc = _ShiftCovers.Find(x => x.Id == c.Id);
                        if (cc != null)
                        {
                            cc.CoverMask = c.CoverMask;
                            cc.CoverName = c.CoverName;
                        }
                    }
                    db.SaveChanges();
                }
            }          
        }

        private bool OnAddCanExecuted(object arg)
        {
            return true;
        }

        private void OnAddExecuted(object obj)
        {
            _dialogService.Show("DetailCoverDialog", OnDetailCallBack);
        }
        private void OnPlanSelected(object obj)
        {
        }
        private void OnSelectionChanged(object obj)
        {
            if(SelectedCalendar.IsLocked == false)
            {
                var ob = obj as object[];
                if (ob != null)
                {
                    int index = (int)ob[0];
                    var oldWeek = SelectedCalendar.ShiftWeeks[index];
                    ShiftWeek newWeek;
                    if (ob[1] is ShiftWeek sw)
                    {                      
                        if (sw.Id == oldWeek.Id) return;
                        newWeek = (ShiftWeek)sw.Clone();
                        newWeek.YearKW = oldWeek.YearKW;
                    }
                    else if (ob[1] is string kw) 
                    {
                        if(oldWeek.YearKW == kw) return;
                        newWeek = oldWeek;
                        newWeek.YearKW = kw;
                    }
                    else
                    {
                        return;
                    }

                    SelectedCalendar.ShiftWeeks[index] = newWeek;
                    using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    var cal = db.ShiftCalendars.Include(x => x.ShiftCalendarShiftPlans).Single(y => y.Id == SelectedCalendar.id);
                    if (cal.ShiftCalendarShiftPlans.Count > index)
                    {
                        var s = cal.ShiftCalendarShiftPlans.ElementAt(index);
                        s.PlanId = newWeek.Id;
                        s.YearKw = newWeek.YearKW;
                    }
                    else
                    {
                        var newWeekdb = new ShiftCalendarShiftPlan() { CalId = SelectedCalendar.id, PlanId = newWeek.Id, YearKw = newWeek.YearKW };
                        cal.ShiftCalendarShiftPlans.Add(newWeekdb);
                    }
                    db.SaveChanges();
                    
                }
            }

        }
        private bool OnDeleteCanExecuted(object arg)
        {
            if (arg is ShiftCover cover)
            {
                return !cover.Lock;
            }
            if(arg is ShiftWeek plan)
            {
                return !plan.Lock;
            }
            return false;
        }

        private void OnDeleteExecuted(object obj)
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            if (obj is ShiftCover cover)
            {
                _ShiftCovers.Remove(cover);
                
                db.Remove(cover);
                db.SaveChanges();
                ShiftCovers.Refresh();
            }
            if(obj is ShiftWeek plan)
            {
                _ShiftWeekPlans.Remove(plan);
                ShiftWeekPlans.Refresh();
                SelectedPlan = _ShiftWeekPlans.First();
                var sp = db.ShiftPlans.Single(x => x.Id == plan.Id);
                db.ShiftPlans.Remove(sp);
                db.SaveChanges();
            }
        }

        private bool OnSaveNewCanExecute(object arg)
        {
            return true;
        }

        private void OnSaveNewExecuted(object obj)
        {
            _dialogService.ShowDialog("InputDialog", OnInputCallback);
        }

        private void OnInputCallback(IDialogResult result)
        {
            if (result.Result == ButtonResult.OK)
            {
                var input = result.Parameters.GetValue<string>("InputText");
                ShiftWeek week = (ShiftWeek)_SelectedPlan.Clone();
                week.Lock = false;
                week.ShiftPlanName = input;

                _ShiftWeekPlans.Add(week);
                ShiftWeekPlans.Refresh();
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var shiftPlan = new ShiftPlan();

                shiftPlan.PlanName = input;
                shiftPlan.Lock = false;
                shiftPlan.Sun = week.GetDayDefinition(0);
                shiftPlan.Mon = week.GetDayDefinition(1);
                shiftPlan.Tue = week.GetDayDefinition(2);
                shiftPlan.Wed = week.GetDayDefinition(3);
                shiftPlan.Thu = week.GetDayDefinition(4);
                shiftPlan.Fre = week.GetDayDefinition(5);
                shiftPlan.Sat = week.GetDayDefinition(6);

                db.ShiftPlans.Add(shiftPlan);

                db.SaveChanges();
                week.Id = shiftPlan.Id;
            }
        }

        private bool OnSaveAllCanExecute(object arg)
        {
            bool result = false;
            if (!PermissionsProvider.GetInstance().GetUserPermission(Permissions.AdminFunc))
            {
                if (SelectedPlan != null)
                {
                    result = !SelectedPlan.Lock;
                    result = shiftCalendars.All(x => x.IsLocked == false);
                    result = shiftCalendars.Any(x => x.IsChanged);
                }
            }
            else result = shiftCalendars.Any(x => x.IsChanged);
            return result;
        }

        private void OnSaveAllExecuted(object obj)
        {
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            if (_SelectedPlan != null)
            {
                
                var planDb = db.ShiftPlans.SingleOrDefault(x => x.Id == _SelectedPlan.Id);
                if (planDb != null)
                {
                    planDb.Sun = _SelectedPlan.GetDayDefinition(0);
                    planDb.Mon = _SelectedPlan.GetDayDefinition(1);
                    planDb.Tue = _SelectedPlan.GetDayDefinition(2);
                    planDb.Wed = _SelectedPlan.GetDayDefinition(3);
                    planDb.Thu = _SelectedPlan.GetDayDefinition(4);
                    planDb.Fre = _SelectedPlan.GetDayDefinition(5);
                    planDb.Sat = _SelectedPlan.GetDayDefinition(6);                  
                }
                else
                {
                    var plan = _SelectedPlan.GetNewShiftPlan();
                    db.ShiftPlans.Add(plan);                 
                }
            }
            foreach (var cal in shiftCalendars.Where(x => x.IsChanged))
            {
                var dbcal = db.ShiftCalendars.Include(x => x.ShiftCalendarShiftPlans).First(y => y.Id == cal.id);
                dbcal.Repeat = cal.Repeat;
                dbcal.CalendarName = cal.CalendarName;
                cal.IsChanged = false;
                for (int i =0; i > cal.ShiftWeeks.Count; i++)
                {
                    var c = dbcal.ShiftCalendarShiftPlans.ElementAtOrDefault(i);
                    if (c != null)
                    {
                        c.PlanId = cal.ShiftWeeks[i].Id;
                        c.YearKw = cal.ShiftWeeks[i].YearKW;
                    }
                    else
                    {
                        dbcal.ShiftCalendarShiftPlans.Add(new ShiftCalendarShiftPlan()
                        { CalId = cal.id, PlanId = cal.ShiftWeeks[i].Id, YearKw = cal.ShiftWeeks[i].YearKW });
                    }
                }
            }
            db.SaveChanges();
        }
        public void DragOver(IDropInfo dropInfo)
        {
            if (PermissionsProvider.GetInstance().GetUserPermission(Permissions.AdminFunc) ||
                _SelectedPlan.Lock == false)
            { 
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.Effects = DragDropEffects.All;
            }           
        }

        public void Drop(IDropInfo dropInfo)
        {
            ShiftCover data = (ShiftCover)dropInfo.Data;
            ShiftDay? item = dropInfo.TargetItem as ShiftDay;

            BitArray cover = new BitArray(data.CoverMask);
            if (IsRubberChecked)
            {
                var s = (ShiftWeek)_SelectedPlan.Clone();
                var d = s.ShiftWeekDays.Single(x => x.Id == item.Id);
                
                d.Definition.And(cover.Not());
                SelectedPlan = null;
                SelectedPlan = s;
            }
            else 
            {                
                var s = (ShiftWeek)_SelectedPlan.Clone();
                var d = s.ShiftWeekDays.Single(x => x.Id == item.Id);
                d.Definition.Or(cover);
                SelectedPlan = null;
                SelectedPlan = s;
            }           
        }
        public class ShiftDay(int id, BitArray definition) : ViewModelBase
        {
            public readonly int Id = id;
            public string WeekDayName { get; } = DateTimeFormatInfo.CurrentInfo.GetDayName((DayOfWeek)id);
            private BitArray _Definition = definition;
            public BitArray Definition {  get { return _Definition; } }
        }
        public interface IShiftWeekProtoType
        {
            int Id { get; set; }
            string ShiftPlanName { get; set; }
            bool Lock {  get; set; }
            string YearKW { get; set; }
            List<ShiftDay> ShiftWeekDays { get; set; }
            object Clone();
            Byte[] GetDayDefinition(int id);
        }
        public class ShiftWeek : IShiftWeekProtoType
        {
            public int Id { get; set; }
            public string ShiftPlanName { get; set; } = string.Empty;
            public bool Lock { get; set; } = false;
            public string YearKW { get; set; } = "200001";

            public List<ShiftDay> ShiftWeekDays { get; set; } = [];

            public ShiftPlan GetNewShiftPlan()
            {
                ShiftPlan shiftPlan = new ShiftPlan();

                shiftPlan.PlanName = ShiftPlanName;
                //shiftPlan.Sun = ShiftWeekDays[0].Definition.CopyTo;

                return shiftPlan;
            }
            public Byte[] GetDayDefinition(int id)
            {
                byte[] bytes = new byte[1440];
                ShiftWeekDays.First(x => x.Id == id).Definition.CopyTo(bytes, 0);

                return bytes;
            }
            public object? Clone(bool deep)
            {
                return (deep) ? Clone() : DeepCopy();
            }

            public object Clone()
            {
                return (ShiftWeek)this.MemberwiseClone();
            }
            // Creates a deep copy
            public object? DeepCopy()
            {
                // use serialized to create a deep copy
                var serialized = JsonSerializer.Serialize(this);
                var copy = JsonSerializer.Deserialize<ShiftWeek>(serialized);
    
                return copy;
            }
        }
        public class ShiftCalendar
        {
            public int id { get; set; }
            public string CalendarName { get; set; } = string.Empty;
            public bool IsLocked { get; set; }
            private bool repeat;
            public bool Repeat
            {
                get { return repeat; }
                set
                {
                    repeat = value;
                    IsChanged = true;
                }
            }
            public bool IsNotRepeat {  get {  return !Repeat; } }
            public bool IsChanged { get; set; } = false;
            public ObservableCollection<ShiftWeek> ShiftWeeks { get; set; } = [];
        }
        void addshift()
        {
            bool[] bo = new bool[1440];
            bool[] nul = new bool[1440];
            bool[] sun = new bool[1440];
            BitArray nulBit = new BitArray(nul);
            byte[] nulByte = new byte[nulBit.Length];

            sun.AsSpan().Slice(1260, 1440 - 1260).Fill(true);
            BitArray sunBit = new BitArray(sun);
            byte[] sunByte = new byte[sunBit.Length];
            sunBit.CopyTo(sunByte, 0);

            //bo.AsSpan().Slice(0, 120).Fill(true);
            //bo.AsSpan().Slice(130, 300 - 130).Fill(true);
            //bo.AsSpan().Slice(1320, 1440 - 1320).Fill(true);

            //bo.AsSpan().Slice(810, 990 - 810).Fill(true);
            //bo.AsSpan().Slice(1000, 1150 - 1000).Fill(true);
            //bo.AsSpan().Slice(1170, 1320 - 1170).Fill(true);

            bo.AsSpan().Slice(120, 15).Fill(true);

            //bo.AsSpan().Slice(300, 510 - 300).Fill(true);
            //bo.AsSpan().Slice(520, 690 - 520).Fill(true);
            //bo.AsSpan().Slice(710, 810 - 710).Fill(true);
            BitArray arr = new BitArray(bo);
            byte[] bytes = new byte[arr.Length];

            arr.CopyTo(bytes, 0);
            var sc = new ShiftPlan()
            {
                PlanName = "3 Schicht",
                Sun = nulByte,
                Mon = bytes,
                Tue = bytes,
                Wed = bytes,
                Thu = bytes,
                Fre = bytes,
                Sat = nulByte,
            };
            var scc = new ShiftCover()
            {
                CoverName = "PausNacht",
                CoverMask = bytes
            };
            using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            db.ShiftCovers.Add(scc);
            db.SaveChanges();
        }
    }
}
