using El2Core.Utils;
using Prism.Events;
using System;
using System.Collections.Generic;
using vhCalendar;

namespace ModuleReport.ViewModels
{
    class SelectionDateViewModel
    {
        public SelectionDateViewModel(IEventAggregator eventAggregator)
        {
            ea = eventAggregator;
            
        }
        IEventAggregator ea;
        RelayCommand? _DateChangedCommand;
        RelayCommand? _DatesChangedCommand;
        public DateTime SelectionDate { get; set; }
        public RelayCommand DateChangedCommand => _DateChangedCommand ??= new RelayCommand(OnDateChanged);
        public RelayCommand DatesChangedCommand => _DatesChangedCommand ??= new RelayCommand(OnDatesChanged);

        private void OnDatesChanged(object obj)
        {
            if (obj is SelectedDatesChangedEventArgs e)
            {
                List<DateTime> dates = [.. e.NewDates];
                ea.GetEvent<MessageReportFilterDateChanged>().Publish(dates);
            }
        }

        private void OnDateChanged(object obj)
        {
            if (obj is SelectedDateChangedEventArgs e)
            {
                List<DateTime> dates = [e.NewDate];
                ea.GetEvent<MessageReportFilterDateChanged>().Publish(dates);
            }
            
        }
    }
}
