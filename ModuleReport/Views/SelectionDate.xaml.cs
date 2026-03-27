using System;
using System.Windows;
using System.Windows.Controls;
using vhCalendar;

namespace ModuleReport.Views
{
    /// <summary>
    /// Interaction logic for SelectionDate.xaml
    /// </summary>
    public partial class SelectionDate : UserControl
    {
        #region ThemePaths
        /// <summary>
        /// Theme Declaration
        /// </summary>
        ThemePath CustomTheme = new ThemePath("CustomTheme", "ModuleReport;component/CustomTheme/CustomTheme.xaml");
        ThemePath BlackMamba = new ThemePath("BlackMamba", "ModuleReport;component/CustomTheme/BlackMamba.xaml");
        #endregion
        public SelectionDate()
        {
            InitializeComponent();
            
        }

        private void UserControl_Initialized(object sender, EventArgs e)
        {
            Cld.RegisterTheme(CustomTheme, typeof(vhCalendar.Calendar));
            Cld.RegisterTheme(BlackMamba, typeof(vhCalendar.Calendar));
            Cld.Theme = CustomTheme.Name;
        }



        private void ButtonSingle_Click(object sender, RoutedEventArgs e)
        {
            Cld.SelectionMode = SelectionType.Single;
            ButtonSingle.SetResourceReference(Control.StyleProperty, "btn_activated");
            ButtonMultiply.SetResourceReference(Control.StyleProperty, "btn_deactivated");
            ButtonWeek.SetResourceReference(Control.StyleProperty, "btn_deactivated");
        }

        private void ButtonMultiply_Click(object sender, RoutedEventArgs e)
        {
            Cld.SelectionMode = SelectionType.Multiple;
            ButtonMultiply.SetResourceReference(Control.StyleProperty, "btn_activated");
            ButtonSingle.SetResourceReference(Control.StyleProperty, "btn_deactivated");
            ButtonWeek.SetResourceReference(Control.StyleProperty, "btn_deactivated");
        }

        private void ButtonWeek_Click(object sender, RoutedEventArgs e)
        {
            Cld.SelectionMode= SelectionType.Week;
            ButtonWeek.SetResourceReference(Control.StyleProperty, "btn_activated");
            ButtonMultiply.SetResourceReference(Control.StyleProperty, "btn_deactivated");
            ButtonSingle.SetResourceReference(Control.StyleProperty, "btn_deactivated");
        }
    }
}
