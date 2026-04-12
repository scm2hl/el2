using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ModuleProxyOrder.Views
{
    /// <summary>
    /// Interaction logic for ProxyOrder.xaml
    /// </summary>
    public partial class ProxyOrderView : UserControl
    {
        public ProxyOrderView()
        {
            InitializeComponent();
        }

        private void ProjectsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListBox lb && lb.SelectedItem != null)
            {
                Proj.IsChecked = false;
            }
        }

        private void CostList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListBox lb && lb.SelectedItem != null)
            {
                Cost.IsChecked = false;
            }
        }

        private void OrdList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListBox lb && lb.SelectedItem != null)
            {
                Ord.IsChecked = false;
            }
        }


    }
}
