using El2Core.Constants;
using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Prism.Commands;
using Prism.Dialogs;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;

namespace ModuleShift.Dialogs
{
    public class DetailCoverVM : ViewModelBase, IDialogAware
    {
        public string Title => "Cover Details";
        public ShiftCover Cover { get; set; }
        public bool IsLocked { get; private set; }
        private ObservableCollection<TimeTuple> TimeList { get; set; } = [];
        public ICollectionView TimeListView { get; private set; }
        ButtonResult result;
        private DelegateCommand? _closeDialogCommand;
        public DelegateCommand CloseDialogCommand =>
            _closeDialogCommand ?? (_closeDialogCommand = new DelegateCommand (OnOkDialog));
        private DelegateCommand? _cancelDialogCommand;
        public DelegateCommand CancelDialogCommand =>
            _cancelDialogCommand ?? (_cancelDialogCommand = new DelegateCommand(OnCancelDialog));
  
        private DelegateCommand? _addNewRowCommand;
        public DelegateCommand AddNewRowCommand =>
            _addNewRowCommand ?? (_addNewRowCommand = new DelegateCommand(OnAddingNewRow, OnCanAddNewRow));

        private bool OnCanAddNewRow()
        {
            return !IsLocked;
        }

        private void OnAddingNewRow()
        {
            TimeList.Add(new TimeTuple());
            
        }

        public DialogCloseListener RequestClose { get; }


        private void OnOkDialog()
        {
            result = ButtonResult.OK;
            bool[] bools = new bool[1440];
            foreach (var t in TimeList)
            {

                if (string.IsNullOrWhiteSpace(t.Start) || string.IsNullOrWhiteSpace(t.End) || string.IsNullOrWhiteSpace(Cover.CoverName)) { break; }
                int start = t.TotalMinute(t.Start);
                int end = t.TotalMinute(t.End);
                if (end == 0) { end = 1440; }
                if (start > end) { break; }
                bools.AsSpan().Slice(start, end - start).Fill(true);
            }
                BitArray bit = new BitArray(bools);
                byte[] bytes = new byte[180];
                bit.CopyTo(bytes, 0);
                Cover.CoverMask = bytes;
       
            OnDialogClosed();
        }
        private void OnCancelDialog()
        {
            result = ButtonResult.Cancel;
            OnDialogClosed();
        }
        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {
            IDialogParameters param = new DialogParameters();

            param.Add("Cover", Cover);
            RequestClose.Invoke(param, result);
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            Cover = (ShiftCover)parameters.GetValue<ShiftCover>("Cover");
            if (Cover == null)
            {
                Cover = new ShiftCover();
                TimeList.Add(new TimeTuple());
                IsLocked = false;
            }
            else
            {
                LoadTimeList();
                IsLocked = Cover.Lock;
                if (IsLocked)
                    IsLocked = !PermissionsProvider.GetInstance().GetUserPermission(Permissions.AdminFunc);
            }
            TimeListView = CollectionViewSource.GetDefaultView(TimeList);
            TimeListView.CollectionChanged += OnTimeListChanged;
            TimeListView.CurrentChanged += OnTimeListCurrentChanged;
         }

        private void OnTimeListCurrentChanged(object? sender, EventArgs e)
        {
            
        }

        private void OnTimeListChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            
        }

        private void LoadTimeList()
        {
            
            bool[] bit = new bool[1440];
            BitArray bitArray = new BitArray(Cover.CoverMask);
            var count = bitArray.Count;
            bitArray.CopyTo(bit, 0);
            int start = 0;
            bool high = false;
            for (int i = 0; i < bit.Length; i++)
            {
                if (bit[i])
                {
                    if (i == 0)
                    {
                        high = true;
                        start = i;
                    }
                    else if (bit[i - 1] == false)
                    {
                        high = true;
                        start = i;
                    }
                    if (i == bit.Length-1)
                    {
                        TimeList.Add(new TimeTuple(start / 60, start % 60, (i / 60), i%60));
                    }
                }

                else if (high)
                {
                    high = false;

                    TimeList.Add(new TimeTuple(start/60, start%60, i/60, i%60));
                }
            }
        }
        public class TimeTuple
        {
            private string? _Start;
            private string? _End;
            public string? Start
            {
                get { return _Start; }
                set { _Start = value; }
            }
            public string? End
            {
                get { return _End; }
                set { _End = value; }
            }
            public TimeTuple() { }
            public TimeTuple(int startHour, int startMinute, int endHour, int endMinute)
            {  _Start = new TimeOnly(startHour, startMinute).ToShortTimeString();
                if(endHour == 0 && endMinute == 0)
                    _End =  "24:00";
                else
                    _End = new TimeOnly(endHour, endMinute).ToShortTimeString();
                
            }

            public int TotalMinute(string timeOnly)
            {
                var dt = DateTime.Parse(timeOnly).TimeOfDay.ToString();
                return (int)Math.Round(TimeSpan.Parse(dt).TotalMinutes);
            }
 
        }

    }
    public class CoverValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            BindingGroup bindingGroup = (BindingGroup)value;
            if(bindingGroup.Items.Count == 2)
            {
                ShiftCover shiftCover = (ShiftCover)bindingGroup.Items[1];
                string coverName = (string)bindingGroup.GetValue(shiftCover, "CoverName");
                DetailCoverVM timeTuple = (DetailCoverVM)bindingGroup.Items[0];
                //string startTime = (string)bindingGroup.GetValue(timeTuple, "Start");
                //string endTime = (string)bindingGroup.GetValue(timeTuple, "End");

                if(string.IsNullOrWhiteSpace(coverName)) return new ValidationResult(false, "Name ist Pflichtfeld");
            }
            return ValidationResult.ValidResult;
        }
    }

}
