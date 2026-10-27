using System;
using System.Collections.Generic;
using System.Text;
using DynamicExpresso;
using El2Core.Utils;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace El2Core.Services
{
    public class CodeSolver
    {
        HolidayLogic holidayLogic = new();
        public void ManipulateWorkPlan(string condition, DateTime date, int startMinute, int endMinute, bool[] workplan, string m)
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
        public (int, int)? ManipulateWorkPlan(string condition, DateTime date, bool[] workplan)
        {

            // Engine direkt im C#-Code starten (die DLLs liegen im ClickOnce-Paket)
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();
            var a = holidayLogic.IsHolyday(date.AddDays(-1));
            var b = holidayLogic.IsHolyday(date);
            var c = holidayLogic.IsHolyday(date.AddDays(1));
            // Ihre Feiertags-Objekte an Python übergeben
            scope.SetVariable("date", date);
            scope.SetVariable("IsHoliday", new Func<DateTime, bool>(holidayLogic.IsHolyday));
            scope.SetVariable("a", a);
            scope.SetVariable("b", b);
            scope.SetVariable("c", c);

// Code des Users ausführen
string userCode = @"
if a and b:
    result = ('00:00', '23:59')
elif b:
    result = ('05:00', '23:59')
elif c and not b:
    result = ('22:00', '23:59')
else:
    result = None
result
";

            dynamic r;
            try
            {
                r = engine.Execute(userCode, scope);
            }
            catch (Exception ex)
            {

                throw new Exception($"Fehler in Regel: {ex.Message}");
            }
            if (r == null) return null;
            else
            { 
                return ((int)TimeSpan.Parse(r[0]).TotalMinutes, (int)TimeSpan.Parse(r[1]).TotalMinutes);
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
