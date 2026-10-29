using El2Core.Models;
using El2Core.ViewModelBase;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using static ModuleShift.ViewModels.ShiftPlanEditViewModel;

namespace ModuleShift.Services
{
    public interface IShiftWeekProtoType
    {
        int Id { get; set; }
        string ShiftPlanName { get; set; }
        bool Lock { get; set; }
        string YearKW { get; set; }
        List<ShiftDay> ShiftWeekDays { get; }
        object Clone();
        Byte[] GetDayDefinition(int id);
    }
    public class ShiftWeek : IShiftWeekProtoType
    {
        public int Id { get; set; }
        public string ShiftPlanName { get; set; } = string.Empty;
        public bool Lock { get; set; } = false;
        public string YearKW { get; set; } = "200001";

        public List<ShiftDay> ShiftWeekDays { get; } = [];

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
    public class ShiftDay(int id, BitArray definition) : ViewModelBase
    {
        public readonly int Id = id;
        public string WeekDayName { get; } = DateTimeFormatInfo.CurrentInfo.GetDayName((DayOfWeek)id);
        private BitArray _Definition = definition;
        public BitArray Definition { get { return _Definition; } }
    }
}
