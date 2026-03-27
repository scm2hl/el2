using System;
using System.Windows;
using System.Windows.Controls;

namespace ModuleReport.Views
{
    /// <summary>
    /// Interaction logic for MaterialResultList.xaml
    /// </summary>
    public partial class MaterialResultList : UserControl
    {
        public MaterialResultList()
        {
            InitializeComponent();
        }

        private void DG_LayoutUpdated(object sender, EventArgs e)
        {
            Thickness t = lblTotal.Margin;

            t.Left = DGMaterials.Columns[0].ActualWidth +
                DGMaterials.Columns[1].ActualWidth +
                DGMaterials.Columns[2].ActualWidth +
                DGMaterials.Columns[3].ActualWidth +
                DGMaterials.Columns[4].ActualWidth;

            lblTotal.Margin = t;
            lblTotal.Width = DGMaterials.Columns[5].ActualWidth;

            lblTotalYieldSum.Width = DGMaterials.Columns[6].ActualWidth;
            lblTotalScrapSum.Width = DGMaterials.Columns[7].ActualWidth;
            lblTotalReworkSum.Width = DGMaterials.Columns[8].ActualWidth;
        }
    }
}
