using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DynamicExpresso;
using El2Core.Utils;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace El2Core.Services
{
    public class CodeSolver
    {
        private static readonly object ScriptLock = new();
        private static readonly ScriptEngine PythonEngine = Python.CreateEngine();
        private static readonly Dictionary<string, ScriptSource> CompiledScripts = new(StringComparer.Ordinal);
        private readonly HolidayLogic holidayLogic = new();


        public (int, int)? ManipulateWorkPlan(DateTime date, string? userCode)
        {
            if (string.IsNullOrWhiteSpace(userCode))
            {
                return null;
            }

            ScriptScope scope = PythonEngine.CreateScope();
            var previousDayHoliday = holidayLogic.IsHolyday(date.AddDays(-1));
            var currentDayHoliday = holidayLogic.IsHolyday(date);
            var nextDayHoliday = holidayLogic.IsHolyday(date.AddDays(1));
            var dayName = date.DayOfWeek.ToString();

            scope.SetVariable("date", date);
            scope.SetVariable("IsHoliday", new Func<DateTime, bool>(holidayLogic.IsHolyday));
            scope.SetVariable("vortag", previousDayHoliday);
            scope.SetVariable("tag", currentDayHoliday);
            scope.SetVariable("folgetag", nextDayHoliday);
            scope.SetVariable("dayName", dayName);

            ScriptSource source;
            lock (ScriptLock)
            {
                if (!CompiledScripts.TryGetValue(userCode, out source))
                {
                    source = PythonEngine.CreateScriptSourceFromString(userCode);
                    CompiledScripts[userCode] = source;
                }
            }

            object? result;
            try
            {
                lock (ScriptLock)
                {
                    result = source.Execute(scope);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Fehler in Regel: {ex.Message}", ex);
            }

            if (result == null)
            {
                return null;
            }

            if (result is IEnumerable enumerable && result is not string)
            {
                var values = enumerable.Cast<object?>()
                    .Select(x => x?.ToString())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToArray();

                if (values.Length == 2 &&
                    TimeSpan.TryParse(values[0], out var start) &&
                    TimeSpan.TryParse(values[1], out var end))
                {
                    return ((int)start.TotalMinutes, (int)end.TotalMinutes);
                }
            }

            return null;
        }
    }
}
