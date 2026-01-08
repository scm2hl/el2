using El2Core.Constants;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;

namespace Lieferliste_WPF.ViewModels
{
    internal class EmployNoteViewModel : ViewModelBase
    {
        public EmployNoteViewModel(IContainerProvider containerProvider, UserSettingsService usrSettingsService,
            IDialogService dialogService)
        {
            container = containerProvider;
            userSettingsService = usrSettingsService;
            _dialogService = dialogService;
            _ctx = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var loggerFactory = container.Resolve<ILoggerFactory>();
            _logger = loggerFactory.CreateLogger<EmployNoteViewModel>();
            VrgTask = new NotifyTaskCompletion<IEnumerable<dynamic>>(LoadVrgAsnc());
            SubmitCommand = new ActionCommand(OnSubmitExecuted, OnSubmitCanExecute);
            ProcessTimeChangeCommand = new ActionCommand(OnProcessTimeChangeExecuted, OnProcessTimeChangeCanExecute);
            LoadingData();
            _dialogService = dialogService;
        }

        public string Title { get; } = "Arbeitszeiten";
        IContainerProvider container;
        UserSettingsService userSettingsService;
        private DB_COS_LIEFERLISTE_SQLContext _ctx;
        private ILogger _logger;
        private readonly IDialogService _dialogService;
        private IEnumerable<dynamic> VorgangRef { get; set; }
        public NotifyTaskCompletion<IEnumerable<dynamic>>? VrgTask { get; private set; }
        private VorgItem? _SelectedVorgangItem;
        public VorgItem? SelectedVorgangItem
        {
            get { return _SelectedVorgangItem; }
            set
            {
                _SelectedVorgangItem = value;
                if (value != null)
                {
                    ReferencePre = new RefItem("Vorgang", value.SourceVorgang.VorgangId, string.Format("{0} - {1}\n{2} {3}",
                        value.Auftrag, value.Vorgang, value.Material?.Trim(), value.Bezeichnung));
 
                }
            }
        }
        private RefItem? _ReferencePre;

        public RefItem? ReferencePre
        {
            get { return _ReferencePre; }
            set
            {
                if (_ReferencePre.Equals(value) == false)
                {
                    _ReferencePre = value;
                    NotifyPropertyChanged(() => ReferencePre);
                }
            }
        }

        private RefItem? _selectedRef;
        public RefItem? SelectedRef
        {
            get { return _selectedRef; }
            set
            {
                if (_selectedRef.Equals(value) == false)
                {
                    _selectedRef = value;
                    ReferencePre = value;                   
                }
            }
        }
        private DateTime _SelectedDate;

        public DateTime SelectedDate
        {
            get { return _SelectedDate; }
            set
            {
                if (_SelectedDate != value)
                {
                    _SelectedDate = value;
                    NotifyPropertyChanged(() => SelectedDate);
                    EmployeeNotesView.Refresh();
                }
            }
        }
        private string quant;

        public string Quant
        {
            get { return quant; }
            set
            {
                quant = value;
                NotifyPropertyChanged(() => Quant);
            }
        }

        private string _Comment = string.Empty;

        public string Comment
        {
            get { return _Comment; }
            set
            {
                _Comment = value;
                NotifyPropertyChanged(() => Comment);
            }
        }
        private double? NoteTime;
        private string _NoteTimePre;

        public string NoteTimePre
        {
            get { return _NoteTimePre; }
            set
            {
                _NoteTimePre = ConvertInputValue(value.ToString(), out NoteTime);
                NotifyPropertyChanged(() => NoteTimePre);
            }
        }

        public List<RefItem> SelectedRefs { get; private set; }
        private string? _SelectedVrgPath;

        public string? SelectedVrgPath
        {
            get { return _SelectedVrgPath; }
            set { _SelectedVrgPath = value; }
        }
        private UserItem _SelectedUser;

        public UserItem SelectedUser
        {
            get { return _SelectedUser; }
            set
            {
                if (_SelectedUser != value)
                {
                    _SelectedUser = value;
                    NotifyPropertyChanged(() => SelectedUser);
                    EmployeeNotesView.Refresh();
                }
            }
        }

        private ObservableCollection<EmployeeNote> EmployeeNotes;
        public ICollectionView EmployeeNotesView { get; private set; }
        public List<string> CalendarWeeks { get; private set; }
        public ICommand SubmitCommand { get; private set; }
        public ICommand ProcessTimeChangeCommand { get; private set; }
        private int _CalendarWeek;

        public int CalendarWeek
        {
            get { return _CalendarWeek; }
            set
            {
                if (_CalendarWeek != value)
                {
                    _CalendarWeek = value;
                    SelectedDate = Get_DateFromKW();
                    
                }
            }
        }

        public List<UserItem> Users { get; private set; }
        private DayOfWeek _SelectedWeekDay;

        public DayOfWeek SelectedWeekDay
        {
            get { return _SelectedWeekDay; }
            set
            {
                _SelectedWeekDay = value;
                SelectedDate = Get_DateFromKW();
            }
        }
        private double _SumTimes;

        public double SumTimes
        {
            get { return _SumTimes; }
            set
            {
                _SumTimes = value;
                NotifyPropertyChanged(() => SumTimes);
            }
        }

        private bool OnProcessTimeChangeCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.EmplCorr);
        }

        private void OnProcessTimeChangeExecuted(object obj)
        {
            if (obj is EmployeeNote empl)
            {
                var par = new DialogParameters();
                par.Add("note", empl);
                _dialogService.ShowDialog("ProcessTimeDialog", par, ProcessTimeChangeCallback);

            }
        }

        private void ProcessTimeChangeCallback(IDialogResult result)
        {
            if (result.Result == ButtonResult.OK)
            {
                var empl = result.Parameters.GetValue<EmployeeNote>("note");
                var corrPre = result.Parameters.GetValue<string?>("newTime");
                _ = ConvertInputValue(corrPre, out double? corr);
                empl.Processingtime = corr;
                empl.RunPropertyChanged();
                SumTimes = EmployeeNotesView.Cast<EmployeeNote>().Sum(x => x.Processingtime ?? 0);
                _logger.LogInformation("{message} {1}", "Set new ProcessTime:", empl.Processingtime);
                _ctx.SaveChanges();
            }
        }
        private async Task<IEnumerable<dynamic>> LoadVrgAsnc()
        {
            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            VorgangRef = await db.Vorgangs
                .Include(x => x.AidNavigation)
                .Include(x => x.AidNavigation.MaterialNavigation)
                .Include(x => x.AidNavigation.DummyMatNavigation)
                .Where(static x => x.AidNavigation.SysStatus.Contains("TABG") == false)                       
                .OrderBy(x => x.Aid)
                .ThenBy(x => x.Vnr)
                .Select(s => new VorgItem(s)).ToListAsync();
            return VorgangRef;
        }

        private void LoadingData()
        {
            var D = DateTime.Today.AddYears(-1);
            EmployeeNotes = _ctx.EmployeeNotes.Where(x => x.Date > D).OrderBy(x => x.Date).ToObservableCollection();

            EmployeeNotesView = CollectionViewSource.GetDefaultView(EmployeeNotes);
            CalendarWeeks = GetKW_List();

            Users = [];
            foreach (var cost in UserInfo.User.AccountCostUnits)
            {
                var us = _ctx.AccountCosts.Include(x => x.Account).Where(x => x.CostId == cost.CostId).AsNoTracking();
                foreach (var account in us)
                {
                    if (Users.All(x => x.User != account.AccountId))
                        Users.Add(new UserItem(account.AccountId, account.Account.Firstname, account.Account.Lastname));
                }               
            }
            SelectedUser = Users.First(x => x.User.Equals(UserInfo.User.UserId, StringComparison.CurrentCultureIgnoreCase));
            EmployeeNotesView.Filter += FilterPredicate;
            EmployeeNotesView.CollectionChanged += CollectionHasChanged;
            var sel = _ctx.EmploySelections.Where(y => y.Active).OrderBy(o => o.Description);
            SelectedRefs = [];
            foreach (var s in sel)
            {
                SelectedRefs.Add(new RefItem("EmploySelection", s.Id.ToString(), s.Description));
            }
            SelectedWeekDay = DateTime.Today.DayOfWeek;

        }

        private void CollectionHasChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Remove)
            {
                foreach (var o in e.OldItems)
                {
                    _ctx.EmployeeNotes.Remove((EmployeeNote)o);
                    _ctx.SaveChanges();
                }
            }

            SumTimes = EmployeeNotesView.Cast<EmployeeNote>().Sum(x => x.Processingtime ?? 0);
        }

        private bool FilterPredicate(object obj)
        {
            if(obj is EmployeeNote note)
            {
                int cw = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(DateTime.Today.AddDays(CalendarWeek*7), CalendarWeekRule.FirstFourDayWeek,
                    DayOfWeek.Sunday);
                return CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(note.Date, CalendarWeekRule.FirstFourDayWeek,
                    DayOfWeek.Sunday) == cw
                    && note.AccId == SelectedUser.User;
            }
            return false;
        }
        private bool OnSubmitCanExecute(object arg)
        {
            return ReferencePre.HasValue;
        }

        private void OnSubmitExecuted(object obj)
        {
            var emp = new EmployeeNote();
            emp.AccId = SelectedUser.User;
            emp.Reference = string.Format("{0}{1}{2}{3}{4}",
                ReferencePre.Value.Table, (char)29, ReferencePre.Value.Id, (char)29, ReferencePre.Value.Description);
            if (ReferencePre.Value.Table == "Vorgang")
            {
                emp.VorgId = ReferencePre.Value.Id;
            }
            else
            {
                emp.SelId = int.Parse(ReferencePre.Value.Id);
            }
            emp.Comment = Comment;
            emp.Stk = Quant;
            emp.Date = SelectedDate;
            emp.Processingtime = NoteTime;
            emp.Usr = UserInfo.User.UserId;
            _ctx.EmployeeNotes.Add(emp);

            _ctx.SaveChanges();
            EmployeeNotes.Add(emp);
            _logger.LogInformation("Employnote Submitted");
            ReferencePre = null;
            Comment = string.Empty;
            Quant = string.Empty;
            NoteTimePre = string.Empty;
        }

        private List<string> GetKW_List()
        {
            List<string> ret = [];
            var date = DateTime.Today.AddDays(-7*(userSettingsService.KWReview-1));

            for (int i = 0; i < userSettingsService.KWReview; i++)
            {
                ret.Add(string.Format("KW {0}", CultureInfo.CurrentCulture.Calendar
                    .GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)));
                date = date.AddDays(7);
            }
            return ret;
        }
        private DateTime Get_DateFromKW()
        {
            var d = DateTime.Today.AddDays(CalendarWeek*7);
            d = d.AddDays((int)SelectedWeekDay - (int)d.DayOfWeek);
            return d;
        }
        private string ConvertInputValue(string? input, out double? noteTime)
        {
            noteTime = default;
            if (input == null) { return string.Empty; }
            int hour = 0, minute = 0;
            bool error = false;
            input = input.Trim();
            //input = input.Replace(',', '.');
            Regex reg = new Regex(@"(\d+):(\d+)");
            Match test = reg.Match(input);
            if (double.TryParse(input, out double t))
            {
                noteTime = t;
                hour = (int)t;
                var m = t - Math.Truncate(t);
                m = Math.Round(m, 2)*60;
                minute = (int)m;
            }
            else if(test.Success)
            {
                if(test.Groups.Count == 3)
                {
                    hour = int.Parse(test.Groups[1].Value);
                    minute = int.Parse((test.Groups[2].Value));
                }
            }
            else
            {
                reg = new Regex(@"^(\d+,?\d*)(\s*[a-zA-Z]+)?");
                test = reg.Match(input);
                
                for(int i = 0; i<2; i++)
                {
                    if (test.Success)
                    {
                        
                        if (test.Groups.Count == 3)
                        {
                            var sec = test.Groups[2].Value.Trim();

                            if (sec.StartsWith("s", StringComparison.CurrentCultureIgnoreCase))
                            {
                                if (double.TryParse(test.Groups[1].Value, out double d))
                                {
                                    noteTime = d;
                                    hour = (int)d;
                                    var m = d - Math.Truncate(d);
                                    m = Math.Round(m, 2) * 60;
                                    minute = (int)m;
                                }
                                    
                            }
                            if (sec.StartsWith("m", StringComparison.CurrentCultureIgnoreCase))
                            {
                                hour += int.Parse(test.Groups[1].Value)/60;
                                minute += int.Parse(test.Groups[1].Value)%60;
                            }
                        }
                    }
                    else { error = true; break; }
                    if (test.Value.Length < input.Length) { test = reg.Match(input[test.Value.Length..]); } else break;
                }
            }
            if (!error)
            {
                if (test.Success)
                    noteTime = hour + Convert.ToDouble(minute) / 60;
                return string.Format("{0}:{1}", hour.ToString(), minute.ToString("D2"));           
            }
            return input;
        }
        public void Closing()
        {         
            _ctx.SaveChanges();
        }
    }

    public class UserItem(string User, string Vorname, string Nachname)
    {
        public string User { get; } = User;
        public string Vorname { get; } = Vorname;
        public string Nachname { get; } = Nachname;
    }
    public struct RefItem(string Table, string Id, string Description)
    {
        public string Table { get; } = Table;
        public string Id { get; } = Id;
        public string Description { get; set; } = Description;
    }
}
