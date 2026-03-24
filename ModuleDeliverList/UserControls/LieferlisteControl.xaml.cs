
using El2Core.Utils;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;

namespace ModuleDeliverList.UserControls
{
    /// <summary>
    /// Interaction logic for LieferlisteView1.xaml
    /// </summary>
    public partial class LieferlisteControl : UserControl
    {
        public LieferlisteControl()
        {
            InitializeComponent();
        }

        #region Properties

        public static readonly DependencyProperty AidProperty
            = DependencyProperty.Register("Aid"
                , typeof(string)
                , typeof(LieferlisteControl),
                new PropertyMetadata("", OnPropertyChanged));
        public String Aid
        {
            get { return (string)GetValue(AidProperty); }
            set { SetValue(AidProperty, value); }
        }
        public static readonly DependencyProperty TtnrProperty
            = DependencyProperty.Register("TTNR"
                , typeof(string)
                , typeof(LieferlisteControl),
                new PropertyMetadata("", OnPropertyChanged, OnTrim));

        private static object OnTrim(DependencyObject d, object baseValue)
        {
            var bsv = baseValue as string;
            var lc = (LieferlisteControl)d;
            return (string.IsNullOrEmpty(bsv)) ? lc.GetValue(DummyMatProperty) : bsv.Trim();           
        }

        public string TTNR
        {
            get { return (string)GetValue(TtnrProperty); }
            set { SetValue(TtnrProperty, value); }
        }
        public static readonly DependencyProperty MatTextProperty
            = DependencyProperty.Register("MatText"
                , typeof(string)
                , typeof(LieferlisteControl),
                new PropertyMetadata("", OnPropertyChanged, WhenDummy));

        private static object WhenDummy(DependencyObject d, object baseValue)
        {
            var bsv = baseValue as string;
            var lc = (LieferlisteControl)d;
            return (string.IsNullOrEmpty(bsv)) ? lc.GetValue(DummyTextProperty) : bsv;
        }


        public string DummyMat
        {
            get { return (string)GetValue(DummyMatProperty); }
            set { SetValue(DummyMatProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DummyMat.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DummyMatProperty =
            DependencyProperty.Register("DummyMat"
                , typeof(string)
                , typeof(LieferlisteControl)
                , new PropertyMetadata("DUMMY", OnPropertyChanged));


        public string MatText
        {
            get { return (string)GetValue(MatTextProperty); }
            set { SetValue(MatTextProperty, value); }
        }

        public string DummyText
        {
            get { return (string)GetValue(DummyTextProperty); }
            set { SetValue(DummyTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DummyText.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DummyTextProperty =
            DependencyProperty.Register("DummyText", typeof(string), typeof(LieferlisteControl), new PropertyMetadata("", OnDummyTextChanged));

        private static void OnDummyTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var lc = (LieferlisteControl)d;
            lc.SetValue(MatTextProperty, e.NewValue);
        }

        public static readonly DependencyProperty EndDateProperty
            = DependencyProperty.Register("EndDate"
                , typeof(DateTime)
                , typeof(LieferlisteControl));

        public DateTime EndDate
        {
            get { return (DateTime)GetValue(EndDateProperty); }
            set { SetValue(EndDateProperty, value); }
        }

        public DateTime EckEnd
        {
            get { return (DateTime)GetValue(EckEndProperty); }
            set { SetValue(EckEndProperty, value); }
        }

        // Using a DependencyProperty as the backing store for EckEnd.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty EckEndProperty =
            DependencyProperty.Register("EckEnd", typeof(DateTime), typeof(LieferlisteControl));



        public string LieferTermin
        {
            get { return (string)GetValue(LieferTerminProperty); }
            set { SetValue(LieferTerminProperty, value); }
        }

        // Using a DependencyProperty as the backing store for LieferTermin.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty LieferTerminProperty =
            DependencyProperty.Register("LieferTermin", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));



        public DateTime? Termin
        {
            get { return (DateTime?)GetValue(TerminProperty); }
            set { SetValue(TerminProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Termin.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty TerminProperty =
            DependencyProperty.Register("Termin", typeof(DateTime?)
                , typeof(LieferlisteControl));


        public static readonly DependencyProperty VnrProperty
            = DependencyProperty.Register("Vnr"
                , typeof(int)
                , typeof(LieferlisteControl),
                new PropertyMetadata(0, OnPropertyChanged));
        public int Vnr
        {
            get { return (int)GetValue(VnrProperty); }
            set { SetValue(VnrProperty, value); }
        }
        public static readonly DependencyProperty VgTextProperty
            = DependencyProperty.Register("VgText"
                , typeof(string)
                , typeof(LieferlisteControl),
                new PropertyMetadata("", OnPropertyChanged));
        public string VgText
        {
            get { return (string)GetValue(VgTextProperty); }
            set { SetValue(VgTextProperty, value); }
        }
        public static readonly DependencyProperty QuantityProperty
            = DependencyProperty.Register("Quantity"
                , typeof(int)
                , typeof(LieferlisteControl));
        public int Quantity
        {
            get { return (int)GetValue(QuantityProperty); }
            set { SetValue(QuantityProperty, value); }
        }
        public static readonly DependencyProperty QuantityScrapProperty
            = DependencyProperty.Register("QuantityScrap"
                , typeof(int)
                , typeof(LieferlisteControl));



        public string SysStatus
        {
            get { return (string)GetValue(SysStatusProperty); }
            set { SetValue(SysStatusProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SysStatus.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SysStatusProperty =
            DependencyProperty.Register("SysStatus", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));


        public int QuantityMiss
        {
            get { return (int)GetValue(QuantityMissProperty); }
            set { SetValue(QuantityMissProperty, value); }
        }

        // Using a DependencyProperty as the backing store for QuantityMiss.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty QuantityMissProperty =
            DependencyProperty.Register("QuantityMiss", typeof(int), typeof(LieferlisteControl), new PropertyMetadata(0));


        public int QuantityScrap
        {
            get { return (int)GetValue(QuantityScrapProperty); }
            set { SetValue(QuantityScrapProperty, value); }
        }
        public static readonly DependencyProperty QuantityYieldProperty
            = DependencyProperty.Register("QuantityYield"
                , typeof(int)
                , typeof(LieferlisteControl));
        public int QuantityYield
        {
            get { return (int)GetValue(QuantityYieldProperty); }
            set { SetValue(QuantityYieldProperty, value); }
        }
        public static readonly DependencyProperty QuantityReworkProperty
            = DependencyProperty.Register("QuantityRework"
                , typeof(int)
                , typeof(LieferlisteControl));
        public int QuantityRework
        {
            get { return (int)GetValue(QuantityReworkProperty); }
            set { SetValue(QuantityReworkProperty, value); }
        }


        public string MarkCode
        {
            get { return (string)GetValue(MarkCodeProperty); }
            set { SetValue(MarkCodeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Marker.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty MarkCodeProperty =
            DependencyProperty.Register("MarkCode", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));


        public string Project
        {
            get { return (string)GetValue(ProjectProperty); }
            set { SetValue(ProjectProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Project.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ProjectProperty =
            DependencyProperty.Register("Project", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));



        public string ProjectInfo
        {
            get { return (string)GetValue(ProjectInfoProperty); }
            set { SetValue(ProjectInfoProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ProjectInfo.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ProjectInfoProperty =
            DependencyProperty.Register("ProjectInfo", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));



        public bool ProjectPrio
        {
            get { return (bool)GetValue(ProjectPrioProperty); }
            set { SetValue(ProjectPrioProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ProjectPrio.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ProjectPrioProperty =
            DependencyProperty.Register("ProjectPrio", typeof(bool), typeof(LieferlisteControl), new PropertyMetadata(false));



        public int? ProjectAttachmentCount
        {
            get { return (int?)GetValue(ProjectAttachmentCountProperty); }
            set { SetValue(ProjectAttachmentCountProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ProjectAttachmentCount.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ProjectAttachmentCountProperty =
            DependencyProperty.Register("ProjectAttachmentCount", typeof(int?), typeof(LieferlisteControl), new PropertyMetadata(null, null, OnAttachment));

        private static object? OnAttachment(DependencyObject d, object baseValue)
        {
            var b = baseValue as int?;
            return b == null || b == 0 ? null : b;
        }

    

        public string WorkArea
        {
            get { return (string)GetValue(WorkAreaProperty); }
            set { SetValue(WorkAreaProperty, value); }
        }

        // Using a DependencyProperty as the backing store for WorkArea.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty WorkAreaProperty =
            DependencyProperty.Register("WorkArea", typeof(string), typeof(LieferlisteControl),
                new PropertyMetadata("", OnPropertyChanged));



        public string WorkAreaText
        {
            get { return (string)GetValue(WorkAreaTextProperty); }
            set { SetValue(WorkAreaTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for WorkAreaText.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty WorkAreaTextProperty =
            DependencyProperty.Register("WorkAreaText", typeof(string), typeof(LieferlisteControl), new PropertyMetadata("-"));



        public string MachineText
        {
            get { return (string)GetValue(MachineTextProperty); }
            set { SetValue(MachineTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for MachineText.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty MachineTextProperty =
            DependencyProperty.Register("MachineText", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));



        public string User
        {
            get { return (string)GetValue(UserProperty); }
            set { SetValue(UserProperty, value); }
        }

        // Using a DependencyProperty as the backing store for User.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty UserProperty =
            DependencyProperty.Register("User", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));



        public string Comment_Me
        {
            get { return (string)GetValue(Comment_MeProperty); }
            set { SetValue(Comment_MeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Comment_Me.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty Comment_MeProperty =
            DependencyProperty.Register("Comment_Me", typeof(string),
                typeof(LieferlisteControl),
                new PropertyMetadata(""));


        public String Comment_Te
        {
            get { return (String)GetValue(Comment_TeProperty); }
            set { SetValue(Comment_TeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Comment_Te.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty Comment_TeProperty =
            DependencyProperty.Register("Comment_Te", typeof(String),
                typeof(LieferlisteControl), new PropertyMetadata(""));


        public string Comment_Ma
        {
            get { return (string)GetValue(Comment_MaProperty); }
            set { SetValue(Comment_MaProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Comment_Ma.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty Comment_MaProperty =
            DependencyProperty.Register("Comment_Ma", typeof(string),
                typeof(LieferlisteControl),
                new PropertyMetadata(""));



        public bool Prio
        {
            get { return (bool)GetValue(PrioProperty); }
            set { SetValue(PrioProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Prio.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty PrioProperty =
            DependencyProperty.Register("Prio", typeof(bool), typeof(LieferlisteControl), new PropertyMetadata(false));



        public string PrioText
        {
            get { return (string)GetValue(PrioTextProperty); }
            set { SetValue(PrioTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for PrioText.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty PrioTextProperty =
            DependencyProperty.Register("PrioText", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));


        public bool InVisible
        {
            get { return (bool)GetValue(InVisibleProperty); }
            set { SetValue(InVisibleProperty, value); }
        }

        // Using a DependencyProperty as the backing store for InVisible.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty InVisibleProperty =
            DependencyProperty.Register("InVisible", typeof(bool), typeof(LieferlisteControl), new PropertyMetadata(false));



        public bool IsSelected
        {
            get { return (bool)GetValue(IsSelectedProperty); }
            set { SetValue(IsSelectedProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Invisible.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register("IsSelected", typeof(bool), typeof(LieferlisteControl), new PropertyMetadata(false));


        public bool Doku
        {
            get { return (bool)GetValue(DokuProperty); }
            set { SetValue(DokuProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Doku.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DokuProperty =
            DependencyProperty.Register("Doku", typeof(bool), typeof(LieferlisteControl), new PropertyMetadata(false));



        public bool Enclosed
        {
            get { return (bool)GetValue(EnclosedProperty); }
            set { SetValue(EnclosedProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Archited.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty EnclosedProperty =
            DependencyProperty.Register("Enclosed", typeof(bool), typeof(LieferlisteControl), new PropertyMetadata(false));


        public string SelectedValue
        {
            get { return (string)GetValue(SelectedValueProperty); }
            set { SetValue(SelectedValueProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SelectedValue.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SelectedValueProperty =
            DependencyProperty.Register("SelectedValue", typeof(string), typeof(LieferlisteControl), new PropertyMetadata(""));

        public static readonly DependencyProperty ExplorerCommandProperty = DependencyProperty.Register("ExplorerCommand", typeof(ICommand), typeof(LieferlisteControl));

        public Dictionary<string, object> AvailableItems
        {
            get { return (Dictionary<string, object>)GetValue(AvailableItemsProperty); }
            set { SetValue(AvailableItemsProperty, value); }
        }

        // Using a DependencyProperty as the backing store for AvailableItems.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AvailableItemsProperty =
            DependencyProperty.Register("AvailableItems", typeof(Dictionary<string, object>), typeof(LieferlisteControl));



        public object SelectedBinding
        {
            get { return (object)GetValue(SelectedBindingProperty); }
            set { SetValue(SelectedBindingProperty, value); }
        }

        // Using a DependencyProperty as the backing store for MyProperty.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SelectedBindingProperty =
            DependencyProperty.Register("SelectedBinding", typeof(object), typeof(LieferlisteControl), new PropertyMetadata(null));



        public int? AttachmentCount
        {
            get { return (int?)GetValue(AttachmentCountProperty); }
            set { SetValue(AttachmentCountProperty, value); }
        }

        // Using a DependencyProperty as the backing store for AttachmentCount.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AttachmentCountProperty =
            DependencyProperty.Register("AttachmentCount", typeof(int?), typeof(LieferlisteControl), new PropertyMetadata(null, null, OnAttachment));


        public string HasMouseOver { get; internal set; }
        #endregion

        //public string ToolTip
        //{
        //    get
        //    {
        //        var toolTip = new StringBuilder();
        //        if (_filterCriterias.Count > 0)
        //        {
        //            toolTip.Append("gefiltert:\n");
        //            foreach (KeyValuePair<String, String> c in _filterCriterias)
        //            {
        //                toolTip.Append(c.Key).Append(" = ").Append(c.Value).Append("\n");
        //            }
        //        }
        //        if (_sortField != string.Empty)
        //        {
        //            toolTip.Append("sortiert nach:\n").Append(_sortField).Append(" / ").Append(_sortDirection);
        //        }

        //        return toolTip.ToString();

        //    }
        //}
        private static void OnPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var newValue = e.NewValue;
            object? val;
            LieferlisteControl? lc = d as LieferlisteControl;
            if (lc != null)
            {
                if (lc.AvailableItems == null) lc.AvailableItems = new Dictionary<string, object>();
                if (lc.AvailableItems.TryGetValue(e.Property.Name.ToLower(), out val))
                {
                    lc.AvailableItems[e.Property.Name.ToLower()] = newValue;
                }
                else
                {
                    lc.AvailableItems.Add(e.Property.Name.ToLower(), newValue);
                }
            }

        }



        private void TextBlock_MouseEnter(object sender, MouseEventArgs e)
        {
            if (sender is TextBlock send)
            {

                Storyboard s = (Storyboard)TryFindResource("EnterStoryBoard");
                Storyboard.SetTarget(s, send);
                s.Begin();
                SelectedValue = send.Name;
            }

        }
        private void TextBlock_MouseLeave(object sender, MouseEventArgs e)
        {
            if (sender is TextBlock send)
            {
                Storyboard s = (Storyboard)TryFindResource("ExitStoryBoard");
                Storyboard.SetTarget(s, send);
                s.Begin();
            }

        }
        private void TextBlock_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock send)
            {
                if (send.IsMouseCaptured)
                    send.ReleaseMouseCapture();
            }
        }

        private void CommentControl_CommentChanged(object sender, RoutedEventArgs e)
        {
            if (sender is WpfCustomControlLibrary.CommentControl cc)
            {
                var id = cc.CommentId;
                var txt = cc.CommentString;
                if (!string.IsNullOrEmpty(txt))
                {
                    El2Core.Services.SubscribeType type = id switch
                    {
                        "BemMe" => El2Core.Services.SubscribeType.MeBem,
                        "BemTe" => El2Core.Services.SubscribeType.TeBem,
                        "BemMa" => El2Core.Services.SubscribeType.MaBem,
                        _ => default
                    };

                    if (!Equals(type, default(El2Core.Services.SubscribeType)))
                    {
                        _ = Globals.NotifyBroker.SendMessageAsync($"{txt}", type);
                    }
                }
            }
        }
    }


}
