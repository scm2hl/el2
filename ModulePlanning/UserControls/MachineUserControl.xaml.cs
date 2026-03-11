using El2Core.Services;
using ModulePlanning.Planning;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ModulePlanning.UserControls
{
    /// <summary>
    /// Interaction logic for MachineUserControl.xaml
    /// </summary>
    public partial class MachineUserControl : UserControl
    {
        public MachineUserControl()
        {         
            InitializeComponent();
        }

        private void Pl_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if(e.PropertyName == "ScrollItem")
            {
                var pl = FindName("Planed1") as DataGrid;

                if(pl == null) { pl = FindName("Planed2") as DataGrid; }
                if(pl != null) { pl.ScrollIntoView((sender as PlanMachine)?.ScrollItem); }
            }
        }

        private void UserControl_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var dtx = this.DataContext as PlanMachine;
            if (dtx != null)
            {
                var list = dtx.ProcessesCV as ListCollectionView;
                if (list.IsAddingNew) { list.CommitNew(); }
                if (list.IsEditingItem) { list.CommitEdit(); }
                
                //dtx.ProcessesCV.SortDescriptions.Clear();
                //dtx.ProcessesCV.SortDescriptions.Add(new SortDescription("SortPos", ListSortDirection.Ascending));
            }
        }

        private void HideDetails_Click(object sender, RoutedEventArgs e)
        {
            //var pl = FindName("Planed") as DataGrid;
            //pl.SelectedIndex = -1;
            //e.Handled = true;
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            var dt = DataContext as PlanMachine;
            this.Resources["Setup"] = Resources[dt.Setup];
    
        }


    }
}
