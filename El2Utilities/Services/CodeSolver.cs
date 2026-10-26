using System;
using System.Collections.Generic;
using System.Text;
using DynamicExpresso;
using El2Core.Utils;

namespace El2Core.Services
{
    public class CodeSolver
    {

        public void ManipulateWorkPlan(string condition, DateTime date, int startMinute, int endMinute, bool[] workplan)
        {
            if (startMinute < 0 ||
                endMinute > workplan.Length ||
                endMinute <= startMinute)
            {
                throw new ArgumentOutOfRangeException();
            }
            RuleContext context = new() { Date = date };
            condition = "IsHoliday(AddDays(date, 1)) && IsMonday(date)";

            Interpreter interpreter = new();
            interpreter.SetVariable("date", context.Date);
            interpreter.SetVariable("dayOfWeek", context.Date.DayOfWeek);
            interpreter.SetFunction("IsHoliday", new Func<DateTime, bool>(context.IsHoliday));
            interpreter.SetFunction("IsSunday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Sunday));
            interpreter.SetFunction("IsMonday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Monday));
            interpreter.SetFunction("IsTuesday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Tuesday));
            interpreter.SetFunction("IsWednesday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Wednesday));
            interpreter.SetFunction("IsThursday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Thursday));
            interpreter.SetFunction("IsFriday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Friday));
            interpreter.SetFunction("IsSaturday", new Func<DateTime, bool>(d => d.DayOfWeek == DayOfWeek.Saturday));
            interpreter.SetFunction("AddDays", new Func<DateTime, int, DateTime>((d, days) => d.AddDays(days)));
            if (interpreter.Eval<bool>(condition))
            {
  
                workplan.AsSpan(startMinute, endMinute - startMinute).Fill(false);
            }


        }
        private class RuleContext
        {
            readonly HolidayLogic holidayLogic = new();
            public DateTime Date { get; set; }
            public bool IsHoliday(DateTime date)
            {
                
                return holidayLogic.IsHolyday(date);
            }

        }
    }
}
