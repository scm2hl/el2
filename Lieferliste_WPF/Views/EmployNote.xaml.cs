using Lieferliste_WPF.ViewModels;
using System.Windows;
using System.Windows.Controls;
using WpfCustomControlLibrary;

namespace Lieferliste_WPF.Views
{
    /// <summary>
    /// Interaction logic for EmployNote.xaml
    /// </summary>
    public partial class EmployNote : UserControl
    {
        public EmployNote()
        {
            InitializeComponent();
        }

        private void UserControl_Unloaded(object sender, RoutedEventArgs e)
        {
            var dt = DataContext as EmployNoteViewModel;
            dt.Closing();
        }

        private void Combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Combo = e.Source as FrameworkElement;
            if (Combo.Name == "RefCombo")
            {
                var cb = this.FindName("VrgCombo") as SearchableComboBox;
                cb.SelectedIndex = -1;
                var prx = this.FindName("PrxCombo") as SearchableComboBox;
                prx.SelectedIndex = -1;
            }
            else if (Combo.Name == "VrgCombo")
            {
                var cb = this.FindName("RefCombo") as ComboBox;
                cb.SelectedIndex = -1;
                var prx = this.FindName("PrxCombo") as SearchableComboBox;
                prx.SelectedIndex = -1;
            }
            else if (Combo.Name == "PrxCombo")
            {
                var cb = this.FindName("RefCombo") as ComboBox;
                cb.SelectedIndex = -1;
                var vrg = this.FindName("VrgCombo") as SearchableComboBox;
                vrg.SelectedIndex = -1;
            }
        }
    }
}
