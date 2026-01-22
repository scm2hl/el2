using System;
using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;


namespace WpfCustomControlLibrary
{
    public class BitPicture1D : ContentControl
    {
        static BitPicture1D()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(BitPicture1D), new FrameworkPropertyMetadata(typeof(BitPicture1D)));

            
        }
        private Canvas? _Canvas { get; set; }
        public SolidColorBrush LowColor
        {
            get { return (SolidColorBrush)GetValue(LowColorProperty); }
            set { SetValue(LowColorProperty, value); }
        }

        // Using a DependencyProperty as the backing store for LowColor.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty LowColorProperty =
            DependencyProperty.Register("LowColor", typeof(SolidColorBrush), typeof(BitPicture1D), new PropertyMetadata(Brushes.White));



        public SolidColorBrush HighColor
        {
            get { return (SolidColorBrush)GetValue(HighColorProperty); }
            set { SetValue(HighColorProperty, value); }
        }

        // Using a DependencyProperty as the backing store for HighColor.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty HighColorProperty =
            DependencyProperty.Register("HighColor", typeof(SolidColorBrush), typeof(BitPicture1D), new PropertyMetadata(Brushes.Peru));



        public double StripeHeight
        {
            get { return (double)GetValue(StripeHeightProperty); }
            set { SetValue(StripeHeightProperty, value); }
        }

        // Using a DependencyProperty as the backing store for StripeHeight.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty StripeHeightProperty =
            DependencyProperty.Register("StripeHeight", typeof(double), typeof(BitPicture1D), new PropertyMetadata(1.0));



        public string WeekDay
        {
            get { return (string)GetValue(WeekDayProperty); }
            set { SetValue(WeekDayProperty, value); }
        }

        // Using a DependencyProperty as the backing store for WeekDay.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty WeekDayProperty =
            DependencyProperty.Register("WeekDay", typeof(string), typeof(BitPicture1D), new PropertyMetadata(""));



        public string WorkRangePattern
        {
            get { return (string)GetValue(WorkRangePatternProperty); }
            set { SetValue(WorkRangePatternProperty, value); }
        }

        // Using a DependencyProperty as the backing store for WorkRangePattern.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty WorkRangePatternProperty =
            DependencyProperty.Register("WorkRangePattern", typeof(string), typeof(BitPicture1D), new PropertyMetadata());


        public BitArray BoolPattern
        {
            get { return (BitArray)GetValue(BoolPatternProperty); }
            set { SetValue(BoolPatternProperty, value); }
        }

        // Using a DependencyProperty as the backing store for BoolPattern.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BoolPatternProperty =
            DependencyProperty.Register("BoolPattern", typeof(BitArray), typeof(BitPicture1D), new PropertyMetadata(null));


        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            if (this.Template != null)
            {
                _Canvas = Template.FindName("PART_DayPicture", this) as Canvas;
                _Canvas.Children.Capacity = 1440;
                CreateBitPicture();
            }
        }

        private void CreateBitPicture()
        {
            double scale = _Canvas.Width / 1440.0;
            if (BoolPattern != null)
            {
                bool[] bo = new bool[1440];
                BoolPattern.CopyTo(bo, 0);
                _Canvas.Children.Clear();
                double left = 0;
                int index = 0;
                foreach (var item in bo)
                {
                    if (index % 60 == 0)
                    {
                        TextBlock nr = new()
                        {
                            Text = string.Format("{0}", index / 60),
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Height = StripeHeight / 4
                        };
                        Rectangle l = new() { Height = StripeHeight / 4, Width = 2, Fill = Brushes.Black };
                        Canvas.SetLeft(nr, left);
                        Canvas.SetTop(nr, 0);
                        _Canvas.Children.Add(nr);
                        Canvas.SetLeft(l, left);
                        Canvas.SetTop(l, StripeHeight / 4);
                        _Canvas.Children.Add(l);
                    }

                    if (item)
                    {
                        Rectangle high = new() { Width = scale, Height = StripeHeight, Fill = HighColor };
                        Canvas.SetLeft(high, left);
                        Canvas.SetTop(high, StripeHeight / 2);
                        _Canvas.Children.Add(high);
                    }
                    else
                    {
                        Rectangle low = new() { Width = scale, Height = StripeHeight, Fill = LowColor };
                        Canvas.SetLeft(low, left);
                        Canvas.SetTop(low, StripeHeight / 2);
                        _Canvas.Children.Add(low);
                    }
                    left += scale;
                    index++;
                }
            }
        }

        public class MinuteLine : Shape
        {
            double LineHeight { get; set; }
            double LineWidth { get; set; }
            private Line line;
            public Line Line => line;

            protected override Geometry DefiningGeometry => throw new NotImplementedException();

            public MinuteLine()
            {
                _ = new MinuteLine(5.0, 0.3);

            }
            public MinuteLine(double lineHeight, double lineWidth)
            {
                LineHeight = lineHeight;
                LineWidth = lineWidth;

                line = new Line();
                line.Height = LineHeight;
                line.Width = LineWidth;
                line.Fill = Brushes.White;
            }
        }
    }
}
