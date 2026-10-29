using El2Core.Models;
using El2Core.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ModuleShift.Services;
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

namespace ModuleShift.Dialogs
{
    public class WeekPreviewViewModel : IDialogAware
    {
        public string Title { get; } = "Wochenvorschau";
        private IContainerExtension _container;
        private ILogger logger;
        public ObservableCollection<WorkShiftService> WorkShiftCollection { get; private set; } = [];
        public ObservableCollection<DataTable> ShiftDataSets { get; private set; } = [];
        public ICollectionView ShiftView { get; private set; }
        public ICollectionView ShiftItemsView { get; private set; }

        private List<ShiftWeek> ShiftWeeks { get; } = [];
        public ShiftWeek? SelectedWeek { get; private set; }

        public DialogCloseListener RequestClose { get; }

        public WeekPreviewViewModel(IContainerExtension container)
        {
            _container = container;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<WeekPreviewViewModel>();
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

                var calendar = db.ShiftCalendars
                    .Include(x => x.ShiftCalendarShiftPlans)
                    .ThenInclude(x => x.Plan)
                    .Where(x => x.Id==4)
                    .FirstOrDefault();

                var StartDay = DateTime.Now.AddDays(-14);
                StartDay = StartDay.AddDays(-(int)StartDay.DayOfWeek);
                var CodeSolver = new CodeSolver();
                ShiftDay sd; bool[] dayplan = new bool[1440];
                for (DateTime i = StartDay; i < StartDay.AddDays(28); i = i.AddDays(7))
                {
                    (int, int)? cs = null;
                    foreach (var w in calendar.ShiftCalendarShiftPlans)
                    {

                        var week = new ShiftWeek
                        {
                            ShiftPlanName = w.Plan.PlanName,
                            YearKW = string.Format("{0} KW {1}", i.ToString("yyyy"), CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(i, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString())
                        };
                        Byte[] bytes;
                        bytes = w.Plan.Sun;
                        sd = new ShiftDay(0, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i);
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);
                        bytes = w.Plan.Mon;
                        sd = new ShiftDay(1, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i.AddDays(1));
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);
                        bytes = w.Plan.Tue;
                        sd = new ShiftDay(2, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i.AddDays(2));
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);
                        bytes = w.Plan.Wed;
                        sd = new ShiftDay(3, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i.AddDays(3));
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);
                        bytes = w.Plan.Thu;
                        sd = new ShiftDay(4, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i.AddDays(4));
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);
                        bytes = w.Plan.Fre;
                        sd = new ShiftDay(5, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i.AddDays(5));
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);
                        bytes = w.Plan.Sat;
                        sd = new ShiftDay(6, new BitArray(bytes));
                        cs = CodeSolver.ManipulateWorkPlan(i.AddDays(6));
                        if (cs != null)
                        {
                            sd.Definition.CopyTo(dayplan, 0);
                            dayplan.AsSpan((int)cs.Value.Item1, cs.Value.Item2 - cs.Value.Item1).Fill(false);
                            sd.Definition.And(new BitArray(dayplan));
                        }
                        week.ShiftWeekDays.Add(sd);

                        ShiftWeeks.Add(week);

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
            
        }
        public class ShiftTimeLine
        {
            public DateTime StartDate { get; set; } = DateTime.Now;
            public DateTime EndDate { get; set; }
            public int WeekNumber { get { return CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(StartDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday); } }
        }
    }
}
