using El2Core.Utils;
using El2Core.ViewModelBase;
using LiveCharts;
using LiveCharts.Wpf;
using ModuleReport.ReportSources;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace ModuleReport.ViewModels
{
    internal class MaterialResultChartViewModel : ViewModelBase
    {
        public MaterialResultChartViewModel(IMaterialSource materialSource, IEventAggregator eventAggregator)
        {
            Materials = materialSource.Materials;
            Materials.CollectionChanged += Materials_CollectionChanged;
            ea = eventAggregator;
            ea.GetEvent<MessageReportFilterDateChanged>().Subscribe(OnDateChanged);
            ea.GetEvent<MessageReportFilterWorkAreaChanged>().Subscribe(OnWorkAreaChanged);
            ea.GetEvent<MessageReportTextSearch>().Subscribe(OnTextSearch);
            SeriesCollection = new SeriesCollection
            {
                new ColumnSeries { Title = "Gut", Values = new ChartValues<int>() },
                new ColumnSeries { Title = "Ausschuss", Values = new ChartValues<int>() },
                new ColumnSeries { Title = "Nacharbeit", Values = new ChartValues<int>() }
            };
            Formatter = value => value.ToString("N");
        }


        IEventAggregator ea;
        private ObservableCollection<ReportMaterial> Materials;
        public SeriesCollection SeriesCollection { get; set; }
        private string textSearch;
        private string[] labels = [];
        public string[] Labels
        {
            get { return labels; }
            set 
            {
                if (labels != value)
                {
                    labels = value;
                    NotifyPropertyChanged(() => Labels);
                }
            }
        }
        private HashSet<int> FilterRids = [];
        private List<DateTime> Dates = [];
        public Func<double, string> Formatter { get; set; }
        private void Materials_CollectionChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            RefreshData();
        }

        private void OnTextSearch(string obj)
        {
            textSearch = obj;
            RefreshData();
        }
        private void OnWorkAreaChanged((int, bool) tuple)
        {
            if(tuple.Item2)
                FilterRids.Add(tuple.Item1);
            else FilterRids.Remove(tuple.Item1);

            RefreshData();
        }

        private void OnDateChanged(List<DateTime> times)
        {
            Dates = times;
            RefreshData();
        }
        private void RefreshData()
        {
            SeriesCollection[0].Values.Clear();
            SeriesCollection[1].Values.Clear();
            SeriesCollection[2].Values.Clear();
            HashSet<string> keys = [];
            int yield = 0, scrap = 0, rework = 0;
            foreach (var date in Dates)
            {

                foreach (var item in Materials.Where(x => x.Date_Time.Date == date).GroupBy(x => x.Rid))
                {
                    if (FilterRids.Any(x => x == item.Key))
                    {
                        Regex? regex;
                        foreach (var mat in item)
                        {
                            if (string.IsNullOrWhiteSpace(textSearch) == false)
                            {
                                regex = new Regex(textSearch, RegexOptions.IgnoreCase);
                                if (regex.Match(mat.TTNR).Success)
                                {
                                    yield += mat.Yield;
                                    scrap += mat.Scrap;
                                    rework += mat.Rework;
                                    keys.Add(mat.MachName);
                                }
                            }
                            else
                            {
                                yield += mat.Yield;
                                scrap += mat.Scrap;
                                rework += mat.Rework;
                                keys.Add(mat.MachName);

                            }
                        }
                        if (yield != 0 || scrap != 0 || rework != 0)
                        {
                            SeriesCollection[0].Values.Add(yield);
                            SeriesCollection[1].Values.Add(scrap);
                            SeriesCollection[2].Values.Add(rework);
                            yield = 0; scrap = 0; rework = 0;
                        }
                    }
                }
            }
            Labels = keys.ToArray();
            
        }
    }
}
