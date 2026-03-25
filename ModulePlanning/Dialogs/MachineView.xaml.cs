using El2Core.Converters;
using El2Core.Models;
using El2Core.Utils;
using ModulePlanning.Dialogs.ViewModels;
using ModulePlanning.Planning;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;

namespace ModulePlanning.Dialogs
{
    /// <summary>
    /// Interaction logic for MachineView.xaml
    /// </summary>
    public partial class MachineView : UserControl
    {

        public MachineView()
        {
            InitializeComponent();
        }


        private void Process_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == "CostId" || e.PropertyName == "WorkArId" || e.PropertyName == "resv" || e.PropertyName == "Bullet")
            {
                e.Cancel = true;
            }
            else if (e.PropertyName == "Aid") e.Column.Header = "Auftrg";
            else if (e.PropertyName == "QuantityMiss") e.Column.Header = "offene Menge";
            else if (e.PropertyName == "Text") e.Column.Header = "Vorgang Kurztext";
            else if (e.PropertyName == "Arbid") e.Column.Header = "Arbeitsplatz";
            else if (e.PropertyType == typeof(DateTime?) || e.PropertyType == typeof(DateTime))
            {
                DataGridTextColumn? dgtc = e.Column as DataGridTextColumn;
                DateConverter con = new();
                if (dgtc != null)
                    ((Binding)dgtc.Binding).Converter = con;
            }
        }


        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var dp = sender as DatePicker;
            if (dp?.DataContext is Vorgang vrg) { vrg.Termin = dp?.SelectedDate; }
        }

        private void CommentControl_CommentChanged(object sender, RoutedEventArgs e)
        {
            if (sender is WpfCustomControlLibrary.CommentControl cc)
            {
                var id = cc.CommentId;
                var txt = cc.CommentString;
                var pl = (MachineViewVM)this.DataContext;
                var vrg = (Vorgang)pl.PlanMachine.ProcessesCV.CurrentItem;
                var refTxt = string.Join(" - ", vrg.AidNavigation.Material, vrg.AidNavigation.MaterialNavigation?.Bezeichng, vrg.Aid, vrg.Vnr, vrg.Text);
                var msg = string.Join((char)29, txt, refTxt);
                
                _ = Globals.NotifyBroker.SendMessageAsync(string.Format("{0}", msg), El2Core.Services.SubscribeType.TeBem);
            }
        }
    }
}
