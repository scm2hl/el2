using El2Core.Models;
using El2Core.Utils;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ModuleShift.Services;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Data;
using System.Xml;
using static El2Core.Constants.ShiftTypes;
using static MaterialDesignThemes.Wpf.Theme.ToolBar;
using static ModuleShift.ViewModels.ShiftPlanEditViewModel;

namespace ModuleShift.ViewModels
{
    public class WeekPreviewViewModel : IDialogAware
    {
        public string Title { get; } = "Wochenvorschau";
        private IContainerExtension _container;
        private ILogger logger;
        public ObservableCollection<WorkShiftService> WorkShiftCollection { get; private set; } = [];
        public ObservableCollection<DataTable> ShiftDataSets { get; private set; } = [];
        public ICollectionView ShiftView { get; private set; }
        public ICollectionView ShiftItemsView { get; private set; }

        private List<ShiftWeek> ShiftWeeks { get; } = [];
 

        public DialogCloseListener RequestClose { get; }

        public WeekPreviewViewModel(IContainerExtension container)
        {
            _container = container;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<WeekPreviewViewModel>();
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
    
                var calendar = db.ShiftCalendars
                    .Include(x => x.ShiftCalendarShiftPlans)
                    .ThenInclude(x => x.Plan)
                    .FirstOrDefault();

                foreach (var w in calendar.ShiftCalendarShiftPlans)
                {
                                        
                    var week = new ShiftWeek();
                        Byte[] bytes;
                    bytes = w.Plan.Sun;
                    week.ShiftWeekDays.Add(new ShiftDay(0, new BitArray(bytes)));
                    bytes = w.Plan.Mon;
                    week.ShiftWeekDays.Add(new ShiftDay(1, new BitArray(bytes)));
                    bytes = w.Plan.Tue;
                    week.ShiftWeekDays.Add(new ShiftDay(2, new BitArray(bytes)));
                    bytes = w.Plan.Wed;
                    week.ShiftWeekDays.Add(new ShiftDay(3, new BitArray(bytes)));
                    bytes = w.Plan.Thu;
                    week.ShiftWeekDays.Add(new ShiftDay(4, new BitArray(bytes)));
                    bytes = w.Plan.Fre;
                    week.ShiftWeekDays.Add(new ShiftDay(5, new BitArray(bytes)));
                    bytes = w.Plan.Sat;
                    week.ShiftWeekDays.Add(new ShiftDay(6, new BitArray(bytes)));

                    ShiftWeeks.Add(week);
                    

                }
                ShiftView = CollectionViewSource.GetDefaultView(ShiftWeeks);
            }
            catch (System.Exception e)
            {

                logger.LogError("{message}", e.ToString());
            }
        }

        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {
            
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            
        }
    }
}
