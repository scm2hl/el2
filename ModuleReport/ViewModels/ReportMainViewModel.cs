using El2Core.Constants;
using El2Core.Utils;
using El2Core.ViewModelBase;
using ModuleReport.Views;
using Prism.Events;
using Prism.Ioc;
using Prism.Navigation.Regions;
using System.Windows.Input;

namespace ModuleReport.ViewModels
{
    internal class ReportMainViewModel : ViewModelBase
    {
        public ReportMainViewModel(IContainerProvider containerProvider, IRegionManager regionManager, IEventAggregator eventAggregator)
        {
            container = containerProvider;
            _regionManager = regionManager.CreateRegionManager();
            ea = eventAggregator;
            
 
            _regionManager.RegisterViewWithRegion<MaterialResultList>(RegionNames.ReportViewRegion);
            _regionManager.RegisterViewWithRegion<SelectionWorkArea>(RegionNames.ReportFilterRegion);
            _regionManager.RegisterViewWithRegion<SelectionDate>(RegionNames.ReportToolRegion);
            _regionManager.RegisterViewWithRegion<MaterialResultChart>(RegionNames.ReportViewRegion1);

            ChangeSourceCommand = new ActionCommand(OnChangeSourceExecuted, OnChangeSourceCanExecute);
            
        }

 

        public string Title { get; } = "Bericht und Auswertungen";

        private IRegionManager _regionManager;
        IContainerProvider container;
        IEventAggregator ea;
        public ICommand ChangeSourceCommand { get; private set; }
        private RelayCommand? _searchCommand;
        public RelayCommand SearchCommand => _searchCommand ??= new RelayCommand(OnTextSearch);

        private void OnTextSearch(object obj)
        {
            if (obj is string text)
            {
                ea.GetEvent<MessageReportTextSearch>().Publish(text);
            }
        }

        private bool OnChangeSourceCanExecute(object arg)
        {
            return true;
        }

        private void OnChangeSourceExecuted(object obj)
        {
            if (int.TryParse((string?)obj, out int nr))
            {
                ea.GetEvent<MessageReportChangeSource>().Publish(nr);
            }
        }
    }
}
