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
        private ShiftPlanEditViewModel.ShiftCalendar? _selectedCalendar { get; set; }
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

                // Code des Users ausführen
                string userCode = @"
if vortag and tag:
    result = ('00:00', '23:59')
elif (not vortag) and tag and dayName != 'Monday':
    result = ('05:00', '23:59')
elif (not tag) and vortag and dayName != 'Tuesday':
    result = ('00:00', '05:00')
elif dayName == 'Sunday' and folgetag:
    result = ('00:00', '23:59')
elif dayName == 'Monday' and tag:
    result = ('00:00', '22:00')
else:
    result = None
result
";
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
                        cs = CodeSolver.ManipulateWorkPlan(dt, userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[0].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[0].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[0].HolidayName = HolidayLogic.GetHolydayName(dt);
                        }
                        cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(1), userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[1].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[1].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[1].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(1));
                        }
  
                        cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(2), userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[2].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[2].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[2].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(2));
                        }
   
                        cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(3), userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[3].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[3].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[3].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(3));
                        }

                        cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(4), userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[4].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[4].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[4].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(4));
                        }
  
                        cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(5), userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[5].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[5].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[5].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(5));
                        }

                        cs = CodeSolver.ManipulateWorkPlan(dt.AddDays(6), userCode);
                        if (cs != null)
                        {
                            w.ShiftWeekDays[6].Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            w.ShiftWeekDays[6].Definition.And(new BitArray(dayplan));
                            w.ShiftWeekDays[6].HolidayName = HolidayLogic.GetHolydayName(dt.AddDays(6));
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
            _selectedCalendar = parameters.GetValue<ShiftPlanEditViewModel.ShiftCalendar>("Calendar");
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
