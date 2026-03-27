using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using static El2Core.Constants.ShiftTypes;


namespace ModuleShift.Services
{
    public interface IProcessStripeService
    { }
    public class ProcessStripeService : IProcessStripeService, IDisposable
    {

        HolidayLogic holidayLogic;
        public ProcessStripeService(IContainerExtension container) { holidayLogic = container.Resolve<HolidayLogic>(); }

        public void Dispose()
        {
        }

        public DateTime GetProcessLength(Vorgang vorgang, DateTime start, out TimeSpan ProcessLength)
        {
            var r = vorgang.Rstze == null ? 0.0 : (short)vorgang.Rstze; //Setup time
            var c = vorgang.Correction == null ? 0.0 : (short)vorgang.Correction; //correction time
            if (vorgang.AidNavigation.Quantity != null && vorgang.AidNavigation.Quantity != 0.0) //if have Total quantity
            {
                //calculation of the currently required time
                var procT = vorgang.Beaze == null ? 0.0 : (short)vorgang.Beaze;
                var quant = (short)vorgang.AidNavigation.Quantity;
                var miss = vorgang.QuantityMissNeo == null ? 0.0 : (short)vorgang.QuantityMissNeo;
                var duration = procT / quant * miss + r + c;

                if (duration > 0)
                {
                    TimeSpan t = TimeSpan.FromMinutes(Convert.ToDouble(duration));
                    ProcessLength = GetCalculatedEndDate(t, start, vorgang.RidNavigation, TimeSpan.Zero);
                    return start.AddTicks(ProcessLength.Ticks);
                }
            }
            ProcessLength = TimeSpan.Zero;
            return start;
        }
        private TimeSpan GetCalculatedEndDate(TimeSpan timeSpan, DateTime start, Ressource ressource, TimeSpan length)
        {
            DateTime dateTime = start;
            if (start.DayOfWeek == DayOfWeek.Saturday)
            {
                dateTime = dateTime.AddDays(1).Date;
                length = length.Add(TimeOnly.MaxValue.ToTimeSpan().Add(TimeOnly.MaxValue.ToTimeSpan()));
            }

            TimeSpan rest = timeSpan;

            var shifts = ressource.RessourceWorkshifts;
            if (shifts.Count == 0) { return TimeSpan.Zero; } //no shifts
            var serializer = XmlSerializerHelper.GetSerializer(typeof(List<WorkShiftItem>));
            TimeOnly endPos = TimeOnly.FromDateTime(dateTime);
            foreach (var shift in shifts.OrderBy(x => x.SidNavigation.ShiftType))
            {
                var sh = shift.SidNavigation;
                if (sh.ShiftType > 10 && sh.ShiftType < 20)
                {
                    if (dateTime.DayOfWeek == DayOfWeek.Sunday) { dateTime = dateTime.AddDays(1).Date; length.Add(TimeOnly.MaxValue.ToTimeSpan()); }
                    while (holidayLogic.IsHolyday(dateTime)) { dateTime = dateTime.AddDays(1).Date; length.Add(TimeOnly.MaxValue.ToTimeSpan()); }
                }
                if (sh.ShiftType > 0 && sh.ShiftType < 10)
                {
                    while (holidayLogic.IsHolyday(dateTime.AddDays(1))) { dateTime = dateTime.AddDays(1).Date; length.Add(TimeOnly.MaxValue.ToTimeSpan()); }
                }

                if (rest.TotalMinutes < 0) break;
                var data = shift.SidNavigation.ShiftDef;
                TextReader reader = new StringReader(data);
                var ws = (List<WorkShiftItem>)serializer.Deserialize(reader);

                foreach (var wsItem in ws)
                {
                    var timespst = wsItem.StartTime.Value.ToTimeSpan();
                    var timespen = wsItem.EndTime.Value.ToTimeSpan();
                    if (length == TimeSpan.Zero)  //the first entry
                    {
                        if (dateTime.TimeOfDay < timespst) //we come from ground
                        {
                            length = length.Add(timespen.Subtract(dateTime.TimeOfDay));
                            rest -= timespen.Subtract(timespst);
                            if (rest.TotalMinutes < 0) //we are override
                            {
                                length = length.Add(rest);
                                break;
                            }
                        }
                        else if (dateTime.TimeOfDay < timespen) //we are between
                        {
                            length = length.Add(timespen.Subtract(dateTime.TimeOfDay));
                            rest -= timespen.Subtract(dateTime.TimeOfDay);
                            if (rest.TotalMinutes < 0) //we are override
                            {
                                length = length.Add(rest);
                                break;
                            }
                        }
                        else { continue; } //get next section
                    }
                    else if (dateTime.TimeOfDay < timespen) //we are between of further
                    {
                        length = length.Add(timespen.Subtract(endPos.ToTimeSpan()));
                        rest -= timespen.Subtract(timespst);
                        endPos = TimeOnly.FromTimeSpan(timespen);
                        if (rest.TotalMinutes < 0) //we are override
                        {
                            length = length.Add(rest);
                            break;
                        }
                    }

                }
            }

            if (rest.TotalMinutes > 0) //we needs a next day?
            {
                length = length.Add(TimeOnly.MaxValue.ToTimeSpan().Subtract(endPos.ToTimeSpan()));
                length = GetCalculatedEndDate(rest, dateTime.AddDays(1).Date, ressource, length);
            }
            return length;
        }

    }
    public class WorkShiftService : ViewModelBase
    {
        public WorkShiftService() { }
        private int _id;
        public int id
        {
            get { return _id; }
            set
            {
                _id = value;
                Changed = true;
            }
        }
        public bool Changed { get; set; } = false;
        private string shiftName;

        public string ShiftName
        {
            get
            {
                return shiftName;
            }
            set
            {
                if (value != shiftName)
                {
                    shiftName = value;
                    NotifyPropertyChanged(() => ShiftName);
                }
            }
        }

        private ShiftType shiftType;

        public ShiftType ShiftType
        {
            get
            {
                return shiftType;
            }
            set
            {
                if (value != shiftType)
                {
                    shiftType = value;
                    NotifyPropertyChanged(() => ShiftType);
                }
            }
        }

        public ObservableCollection<WorkShiftItem> Items { get; set; } = [];
    }
    public class WorkShiftItem : ViewModelBase
    {
        public WorkShiftItem() { }
        [XmlIgnore]
        private TimeOnly? _startTime;
        [XmlIgnore]
        public TimeOnly? StartTime
        {
            get { return _startTime; }
            set
            {
                if (_startTime != value)
                {
                    _startTime = value;
                    NotifyPropertyChanged(() => StartTime);
                }
            }
        }

        [XmlIgnore]
        private TimeOnly? _endTime;
        [XmlIgnore]
        public TimeOnly? EndTime
        {
            get { return _endTime; }
            set
            {
                if (_endTime != value)
                {
                    _endTime = value;
                    NotifyPropertyChanged(() => EndTime);
                }
            }
        }
        [XmlIgnore]
        public bool Changed { get; set; } = false;
        public string? StartTimeProxy
        {
            get { return StartTime.ToString(); }
            set
            {
                if (string.IsNullOrEmpty(value) == false)
                    StartTime = TimeOnly.Parse(value);
                else StartTime = null;
            }
        }
        public string? EndTimeProxy
        {
            get { return EndTime.ToString(); }
            set
            {
                if (string.IsNullOrEmpty(value) == false)
                    EndTime = TimeOnly.Parse(value);
                else EndTime = null;
            }
        }
    }
}
