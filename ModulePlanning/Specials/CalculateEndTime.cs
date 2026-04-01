using El2Core.Models;
using El2Core.Utils;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;
using System.Linq;

namespace ModulePlanning.Specials
{
    public interface IShiftPlan
    { }
    public class ShiftPlanService : IShiftPlan
    {
        private ImmutableArray<bool[]> weekPlan;
        private Dictionary<int, ImmutableArray<bool[]>> weekPlans = [];
        IContainerProvider container;
        private readonly int rid;
        private HolidayLogic holidayLogic;
        private List<Stopage> stoppages;
        private bool repeat;
        private ILogger logger;

        public ShiftPlanService(int rid, IContainerProvider container)
        {
            this.rid = rid;

            this.container = container;
            var factory = container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<ShiftPlanService>();
            ReloadStoppage();
            ReloadShiftCalendar();      
            
            holidayLogic = container.Resolve<HolidayLogic>();
        }

        private bool[] GetManipulateMask(bool[] weekplan, DateTime start)
        {
            var stopes = stoppages.Where(x => start < x.Endtime && (start.Date <= x.Endtime.Date && start.Date >= x.Starttime.Date));
            var ret = weekplan.ToArray();
            foreach (var stop in stopes)
            {
                int begin =0, end =0;

                begin = (stop.Starttime >= start) ? Convert.ToInt32(stop.Starttime.TimeOfDay.TotalMinutes) : 0;
                end = (stop.Endtime.Date == start.Date) ? Convert.ToInt32(stop.Endtime.TimeOfDay.TotalMinutes) : 1440;
                ret.AsSpan(begin, end - begin).Fill(false);           
            }
            
            return ret;
        }

        public DateTime GetEndDateTime(double processLength, DateTime start)
        {
            int key = int.Parse(string.Concat(start.Year.ToString(),
                    CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(start, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)));
            if (repeat)
            {
                var count = weekPlans.Count;
                var diff = key - weekPlans.Keys.First();
                var mod = diff % count;
                var k = weekPlans.Keys.ElementAt(mod);
                weekPlan = weekPlans[k];
            }
            else if (weekPlans.TryGetValue(key, out weekPlan) == false) return start;

            if (weekPlan.IsDefaultOrEmpty || processLength == 0) return start;
            bool[] tmpWeekPlan;
            int resultMinute = 0;
            
            TimeSpan time = start.TimeOfDay;
            if (holidayLogic.IsHolyday(start))
            {
                while (holidayLogic.IsHolyday(start.AddDays(1))) { start = start.AddDays(1); } //move the start to => tomorrow is not holiday
            }
            int startDay = (int)start.DayOfWeek;
            tmpWeekPlan = [.. weekPlan[startDay]];
            if (startDay != 0)
            {
                if (holidayLogic.IsHolyday(start)) {tmpWeekPlan.AsSpan(0, 1320).Fill(false);} //clear the shift unless nightshift_2
                if (holidayLogic.IsHolyday(start.AddDays(1))) tmpWeekPlan.AsSpan(1320, 120).Fill(false); //clear the nightshift_2
            }
            else { if (holidayLogic.IsHolyday(start.AddDays(1))) tmpWeekPlan.AsSpan(1260, 180).Fill(false); } //clear the nightshift_sun
            tmpWeekPlan = GetManipulateMask(tmpWeekPlan, start);

            for (int j = (int)time.TotalMinutes; j < tmpWeekPlan.Length; j++)
            {
                if (tmpWeekPlan[j]) --processLength;
                if (processLength <= 0) { resultMinute = j; break; }
            }
            start = start.Date;

            if (processLength > 0)
            {
                start = GetEndDateTime(processLength, start.AddDays(1));
            }
            else
            {
                start = start.AddMinutes(resultMinute);
            }
            
            return start;
        }
        public void ReloadStoppage()
        {
            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            stoppages = db.Stopages.AsNoTracking().Where(x => x.Rid == rid).ToList();
        }
        public void ReloadShiftCalendar()
        {
            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var scal = db.ShiftCalendars
                .Include(x => x.ShiftCalendarShiftPlans)
                .ThenInclude(x => x.Plan)
                .SingleOrDefault(x => x.Ressources.Any(x => x.RessourceId == rid));
            if (scal != null)
            {
                repeat = scal.Repeat;
                weekPlans.Clear();
                foreach (var s in scal.ShiftCalendarShiftPlans.OrderBy(x => x.YearKw))
                {
                    List<bool[]> days = new List<bool[]>();
                    Byte[] bytes;
                    bool[] sbools = new bool[1440];
                    bytes = s.Plan.Sun;
                    BitArray bitArray = new BitArray(bytes);
                    bitArray.CopyTo(sbools, 0);
                    days.Add(sbools);

                    bytes = s.Plan.Mon;
                    bitArray = new BitArray(bytes);
                    bool[] mbools = new bool[1440];
                    bitArray.CopyTo(mbools, 0);
                    days.Add(mbools);

                    bytes = s.Plan.Tue;
                    bitArray = new BitArray(bytes);
                    bool[] tubools = new bool[1440];
                    bitArray.CopyTo(tubools, 0);
                    days.Add(tubools);

                    bytes = s.Plan.Wed;
                    bitArray = new BitArray(bytes);
                    bool[] wbools = new bool[1440];
                    bitArray.CopyTo(wbools, 0);
                    days.Add(wbools);

                    bytes = s.Plan.Thu;
                    bitArray = new BitArray(bytes);
                    bool[] thbools = new bool[1440];
                    bitArray.CopyTo(thbools, 0);
                    days.Add(thbools);

                    bytes = s.Plan.Fre;
                    bitArray = new BitArray(bytes);
                    bool[] fbools = new bool[1440];
                    bitArray.CopyTo(fbools, 0);
                    days.Add(fbools);

                    bytes = s.Plan.Sat;
                    bitArray = new BitArray(bytes);
                    bool[] sabools = new bool[1440];
                    bitArray.CopyTo(sabools, 0);
                    days.Add(sabools);

                    int key = int.Parse(s.YearKw);
                    weekPlans.Add(key, [.. days]);
                }
            }
        }
    }
}
