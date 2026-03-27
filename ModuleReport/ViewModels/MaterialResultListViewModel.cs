using El2Core.Utils;
using El2Core.ViewModelBase;
using ModuleReport.ReportSources;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Data;

namespace ModuleReport.ViewModels
{
    internal class MaterialResultListViewModel : ViewModelBase
    {
        public MaterialResultListViewModel(IMaterialSource materialSource, IEventAggregator eventAggregator)
        { 
            Materials = CollectionViewSource.GetDefaultView(materialSource.Materials);
            MaterialSource = materialSource;
            ea = eventAggregator;
            ea.GetEvent<MessageReportFilterWorkAreaChanged>().Subscribe(OnFilterWorkAreaReceived);
            ea.GetEvent<MessageReportFilterDateChanged>().Subscribe(OnFilterDateReceived);
            ea.GetEvent<MessageReportTextSearch>().Subscribe(OnTextSerarch);
            Materials.GroupDescriptions.Add(new PropertyGroupDescription("TTNR"));
            Materials.Filter += OnFilterPredicate;
        }

        public ICollectionView Materials { get; }

        readonly IEventAggregator ea;
        private readonly HashSet<int> FilterRids = [];
        private List<DateTime> FilterDates = [];
        private int _YieldSum = 0;
        private int _ScrapSum = 0;
        private int _ReworkSum = 0;
        private string _textSearch = string.Empty;
        public IMaterialSource MaterialSource { get; }
        public int YieldSum
        {
            get { return _YieldSum; }
            set
            {
                _YieldSum = value;
                NotifyPropertyChanged(() => YieldSum);
            }
        }
        public int ReworkSum
        {
            get { return _ReworkSum; }
            set
            {
                _ReworkSum = value;
                NotifyPropertyChanged(() => ReworkSum);
            }
        }
        public int ScrapSum
        {
            get { return _ScrapSum; }
            set
            {
                _ScrapSum = value;
                NotifyPropertyChanged(() => ScrapSum);
            }
        }

        private void OnTextSerarch(string obj)
        {
            _textSearch = obj;
            YieldSum = 0;
            ReworkSum = 0;
            ScrapSum = 0;
            Materials.Refresh();
        }
        private void OnFilterDateReceived(List<DateTime> dates)
        {
            FilterDates = dates;
            YieldSum = 0;
            ReworkSum = 0;
            ScrapSum = 0;
            Materials.Refresh();
        }

        private void OnFilterWorkAreaReceived((int, bool) tuple)
        {

            if (tuple.Item2)
                FilterRids.Add(tuple.Item1);
            else
                FilterRids.Remove(tuple.Item1);
            YieldSum = 0;
            ReworkSum = 0;
            ScrapSum = 0;
            Materials.Refresh();
        }

        private bool OnFilterPredicate(object obj)
        {
            bool accept = false;
            if (obj is ReportMaterial m)
            {
                accept = FilterRids.Any(x => x == m.Rid) && FilterDates.Any(y => m.Date_Time.Date == y);
                if (accept)
                {
                    YieldSum += m.Yield;
                    ScrapSum += m.Scrap;
                    ReworkSum += m.Rework;
                }
                if(accept && string.IsNullOrWhiteSpace(_textSearch) == false)
                {
                    Regex regex = new Regex(_textSearch, RegexOptions.IgnoreCase);
                    accept = regex.Match(m.TTNR).Success;
                }
            }
            return accept;
        }

    }
}
