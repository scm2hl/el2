using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ModuleShift.Services;
using ModuleShift.ViewModels;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Input;
using ShiftCalendar = ModuleShift.Services.ShiftCalendar;

namespace ModuleShift.Dialogs
{
    public class WeekPreviewViewModel : ViewModelBase, IDialogAware
    {
        public string Title { get; } = "Wochenvorschau";
        private IContainerExtension _container;
        private ILogger logger;
        public ICommand PreviousWeekCommand { get; }
        public ICommand NextWeekCommand { get; }
        public ObservableCollection<WorkShiftService> WorkShiftCollection { get; private set; } = [];
        public ObservableCollection<DataTable> ShiftDataSets { get; private set; } = [];
        public ICollectionView ShiftView { get; private set; }
        public ICollectionView ShiftItemsView { get; private set; }
        private ShiftCalendar? _selectedCalendar { get; set; }
        private List<ShiftWeek> ShiftWeeks { get; } = [];
        private ShiftWeek? _selectedWeek;
        public ShiftWeek? SelectedWeek
        {
            get { return _selectedWeek; }
            private set
            {
                _selectedWeek = value;
                NotifyPropertyChanged(() => SelectedWeek);
            }
        }

        public DialogCloseListener RequestClose { get; }

        public WeekPreviewViewModel(IContainerExtension container)
        {
            _container = container;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<WeekPreviewViewModel>();
            PreviousWeekCommand = new ActionCommand(OnPreviousWeekExecute, OnPreviousWeekCanExecute);
            NextWeekCommand = new ActionCommand(OnNextWeekExecute, OnNextWeekCanExecute);

        }

        private bool OnNextWeekCanExecute(object arg)
        {
            return true;
        }

        private void OnNextWeekExecute(object obj)
        {
            SelectedWeek = ShiftWeeks[(ShiftWeeks.IndexOf(SelectedWeek) + 1) % ShiftWeeks.Count];
        }

        private bool OnPreviousWeekCanExecute(object arg)
        {
            return true;
        }
        

        private void OnPreviousWeekExecute(object obj)
        {
            SelectedWeek = ShiftWeeks[(ShiftWeeks.IndexOf(SelectedWeek) - 1 + ShiftWeeks.Count) % ShiftWeeks.Count];
        }

        private void LoadData()
        {
            try
            {
                if (_selectedCalendar == null) return;


                ShiftWeek w;
                var StartDay = DateTime.Now.AddDays(-14);
                StartDay = StartDay.AddDays(-(int)StartDay.DayOfWeek);
                var CodeSolver = new CodeSolver();
                bool[] dayplan = new bool[1440];
                HolidayLogic HolidayLogic = new HolidayLogic();
                for (DateTime dt = StartDay; dt < StartDay.AddDays(28); dt = dt.AddDays(7))
                {
                    (int, int)? cs = null;
                    foreach (var week in _selectedCalendar.ShiftWeeks)
                    {
                        w = (ShiftWeek)week.Clone(true);
                        w.YearKW = $"{dt.Year} KW {CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)}";

                        for (int i = 0; i < 7; i++)
                        {
                            cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(i), w.NightRule);
                            w.ShiftWeekDays[i].Definition.CopyTo(dayplan, 0);
                            if (cs != null)
                            {
                                dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            }
                            else if (HolidayLogic.IsHolyday(dt.AddDays(i)))
                            {
                                dayplan.AsSpan(0, 1440).Fill(false);
                            }

                            w.ShiftWeekDays[i].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[i].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(i));   
                        }
                        ShiftWeeks.Add(w);
                    }
                }
                SelectedWeek = ShiftWeeks[2];
            }
            catch (System.Exception e)
            {

                logger.LogError("{message}", e.ToString());
            }
        }

        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {
            
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            _selectedCalendar = parameters.GetValue<ShiftCalendar>("Calendar");
            LoadData();
        }
        public class ShiftTimeLine
        {
            public DateTime StartDate { get; set; } = DateTime.Now;
            public DateTime EndDate { get; set; }
            public int WeekNumber { get { return CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(StartDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday); } }
        }
    }
}
