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
        public void ManipulateWorkPlan(string condition, DateTime date, int startMinute, int endMinute, bool[] workplan)
        {
            if (startMinute < 0 ||
                endMinute > workplan.Length ||
                endMinute <= startMinute)
            {
                throw new ArgumentOutOfRangeException();
            }
            // Engine direkt im C#-Code starten (die DLLs liegen im ClickOnce-Paket)
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();

            // Ihre Feiertags-Objekte an Python übergeben
            scope.SetVariable("date", date);
            scope.SetVariable("IsHoliday", new Func<DateTime, bool>(holidayLogic.IsHolyday));

            // Code des Users ausführen
            string userCode = "if not IsHoliday(date) and date.DayOfWeek == date.DayOfWeek.Tuesday: MethodeA()";
            engine.Execute(userCode, scope);
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
