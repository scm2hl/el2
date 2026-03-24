using MahApps.Metro.Converters;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static System.Net.Mime.MediaTypeNames;

namespace WpfCustomControlLibrary
{
    /// <summary>
    /// Follow steps 1a or 1b and then 2 to use this custom control in a XAML file.
    ///
    /// Step 1a) Using this custom control in a XAML file that exists in the current project.
    /// Add this XmlNamespace attribute to the root element of the markup file where it is 
    /// to be used:
    ///
    ///     xmlns:MyNamespace="clr-namespace:WpfCustomControlLibrary"
    ///
    ///
    /// Step 1b) Using this custom control in a XAML file that exists in a different project.
    /// Add this XmlNamespace attribute to the root element of the markup file where it is 
    /// to be used:
    ///
    ///     xmlns:MyNamespace="clr-namespace:WpfCustomControlLibrary;assembly=WpfCustomControlLibrary"
    ///
    /// You will also need to add a project reference from the project where the XAML file lives
    /// to this project and Rebuild to avoid compilation errors:
    ///
    ///     Right click on the target project in the Solution Explorer and
    ///     "Add Reference"->"Projects"->[Select this project]
    ///
    ///
    /// Step 2)
    /// Go ahead and use your control in the XAML file.
    ///
    ///     <MyNamespace:CustomControl1/>
    ///
    /// </summary>
    public class CommentControl : Control
    {
        private const bool DefaultValue = false;
        private TextBox _textBox;
        private Canvas _canvas;
        private bool _IsInUse = false;
        
        static CommentControl()
        {
           
            DefaultStyleKeyProperty.OverrideMetadata(typeof(CommentControl), new FrameworkPropertyMetadata(typeof(CommentControl)));
        }

        // Routed event raised when the CommentString changes
        public static readonly RoutedEvent CommentChangedEvent = EventManager.RegisterRoutedEvent(
            "CommentChanged", RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(CommentControl));

        // CLR wrapper for the routed event
        public event RoutedEventHandler CommentChanged
        {
            add { AddHandler(CommentChangedEvent, value); }
            remove { RemoveHandler(CommentChangedEvent, value); }
        }

        public string CommentId
        {
            get { return (string)GetValue(CommentIdProperty); }
            set { SetValue(CommentIdProperty, value); }
        }

        // Using a DependencyProperty as the backing store for CommentId.
        public static readonly DependencyProperty CommentIdProperty =
            DependencyProperty.Register("CommentId", typeof(string), typeof(CommentControl), new PropertyMetadata(default(string)));

        public string CommentString
        {
            get { return (string)GetValue(CommentStringProperty); }
            set { SetValue(CommentStringProperty, value); }
        }

        // Using a DependencyProperty as the backing store for CommentString.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty CommentStringProperty =
            DependencyProperty.Register("CommentString", typeof(string), typeof(CommentControl),
                new PropertyMetadata(default, OnCommentStringChanged));


        public new double FontSize
        {   
            get { return (double)GetValue(FontSizeProperty); }
            set { SetValue(FontSizeProperty, value); }
        }
        public static readonly DependencyProperty FontSizeProperty = DependencyProperty.Register(nameof(FontSize),
                typeof(double), typeof(CommentControl),
                new FrameworkPropertyMetadata(11.0,
                    FrameworkPropertyMetadataOptions.AffectsMeasure));

        internal string HeaderText
        {
            get { return (string)GetValue(HeaderTextProperty); }
            set { SetValue(HeaderTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for HeaderText.  This enables animation, styling, binding, etc...
        internal static readonly DependencyProperty HeaderTextProperty =
            DependencyProperty.Register("HeaderText", typeof(string), typeof(CommentControl),
                new PropertyMetadata(default(string)));


        internal string CommentText
        {
            get { return (string)GetValue(CommentTextProperty); }
            set { SetValue(CommentTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for CommentText.  This enables animation, styling, binding, etc...
        internal static readonly DependencyProperty CommentTextProperty =
            DependencyProperty.Register("CommentText", typeof(string), typeof(CommentControl), new PropertyMetadata(default(string)));

        public string User
        {
            get { return (string)GetValue(UserProperty); }
            set { SetValue(UserProperty, value); }
        }

        // Using a DependencyProperty as the backing store for User.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty UserProperty =
            DependencyProperty.Register("User", typeof(string), typeof(CommentControl), new PropertyMetadata(default(string)));



        public bool IsEditable
        {
            get { return (bool)GetValue(IsEditableProperty); }
            set { SetValue(IsEditableProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsEditable.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsEditableProperty =
            DependencyProperty.Register("IsEditable", typeof(bool), typeof(CommentControl),
                new PropertyMetadata(DefaultValue, OnIsEditableChanged));




        public SolidColorBrush BackGround
        {
            get { return (SolidColorBrush)GetValue(BackGroundProperty); }
            set { SetValue(BackGroundProperty, value); }
        }

        // Using a DependencyProperty as the backing store for BackGround.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BackGroundProperty =
            DependencyProperty.Register(nameof(BackGround), typeof(SolidColorBrush), typeof(CommentControl), new PropertyMetadata(Brushes.Transparent));



        public SolidColorBrush HeaderBrush
        {
            get { return (SolidColorBrush)GetValue(HeaderBrushProperty); }
            set { SetValue(HeaderBrushProperty, value); }
        }

        // Using a DependencyProperty as the backing store for HeaderBrush.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty HeaderBrushProperty =
            DependencyProperty.Register(nameof(HeaderBrush), typeof(SolidColorBrush), typeof(CommentControl), new PropertyMetadata(Brushes.Gold));




        private static void OnIsEditableChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var cmt = (CommentControl)d;
            if (cmt._textBox != null)
            {
                cmt._textBox.IsReadOnly = !(bool)e.NewValue;
            }
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            _canvas = (Canvas)Template.FindName("PART_CanvasSize", this);
            _textBox = (TextBox)Template.FindName("PART_TextBox", this);
            _textBox.IsEnabled = true;
            _textBox.IsReadOnly = true;
            _textBox.SizeChanged += txSizeChanged;
            GotFocus += OnGotFocus;
            LostFocus += OnLostFocus;
            
            RaiseEvent(new RoutedEventArgs(LoadedEvent, this));
        }


        private static void OnCommentStringChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            
            var cmt = (CommentControl)d;
            // If the value didn't change, do nothing
            var oldStr = e.OldValue as string;
            var newStr = e.NewValue as string;
            if (string.Equals(oldStr, newStr, StringComparison.Ordinal))
                return;

            if (!cmt._IsInUse)
            {
                if (newStr is string s)
                {
                    var st = s.Split((char)29);
                    if (st.Length == 2)
                    {
                        cmt.HeaderText = st[0];
                        cmt.CommentText = st[1];
                    }
                }
            }

        }

        private void OnGotFocus(object sender, RoutedEventArgs e)
        {
            _IsInUse = true;
        }

        private void OnLostFocus(object sender, RoutedEventArgs e)
        {
            _IsInUse = false;
            if (_textBox.Text.Length != 0)
            {
                if (CommentText != _textBox.Text)
                {
                    
                    HeaderText = $"[{User} - {DateTime.Now.ToShortDateString()}]";
                    CommentText = _textBox.Text ;
                    CommentString = string.Format("{0}{1}{2}", HeaderText, (char)29, _textBox.Text);

                    // Raise routed event to notify listeners that the comment changed
                    RaiseEvent(new RoutedEventArgs(CommentChangedEvent, this));
                }
            }
            else
            {
                HeaderText = string.Empty;
                CommentText = string.Empty;
                CommentString = default;
            }
            
        }
 
        private void txSizeChanged(object sender, SizeChangedEventArgs e)
        {
            var txt = sender as TextBox;
            _canvas.Width = e.NewSize.Width;
            _canvas.Height = 15 + e.NewSize.Height;
            
        }
    }
}
