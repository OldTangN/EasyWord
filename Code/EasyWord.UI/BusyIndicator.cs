using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Shapes;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows;

namespace EasyWord.UI
{
    public class BusyIndicator : UserControl
    {
        private static readonly DependencyProperty IsBusyProperty = DependencyProperty.Register("IsBusy", typeof(bool), typeof(BusyIndicator), new PropertyMetadata(false, new PropertyChangedCallback(IsBusyPropertyChanged)));

        private static void IsBusyPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            BusyIndicator bi = sender as BusyIndicator;
            if (bi != null)
            {
                bi.OnIsBusyPropertyChanged(e.OldValue, e.NewValue);
            }
        }

        public bool IsBusy
        {
            get
            {
                return (bool)this.GetValue(IsBusyProperty);
            }
            set
            {
                this.SetValue(IsBusyProperty, value);
                if (value)
                {
                    this.Visibility = Visibility.Visible;
                }
                else
                {
                    this.Visibility = Visibility.Collapsed;
                }                
            }
        }

        public BusyIndicator()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(BusyIndicator_Loaded);
        }

        void BusyIndicator_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.Parent != null)
            {
                FrameworkElement fe = this.Parent as FrameworkElement;
                //this.Margin = new Thickness((fe.ActualWidth - this.ActualWidth) / 2, (fe.ActualHeight - this.ActualHeight) / 2, 0, 0);
            }
        }

        private void InitializeComponent()
        {
            this.HorizontalAlignment = HorizontalAlignment.Center;
            this.HorizontalContentAlignment = HorizontalAlignment.Center;
            this.VerticalAlignment = VerticalAlignment.Center;
            this.VerticalContentAlignment = VerticalAlignment.Center;
            this.IsBusy = false;
            Grid g = new Grid();          
            //g.Width = 200;
            //g.Height = 200;           
            this.Content = g;
            List<Path> lstPath = new List<Path>();

            double dAnimationTimeSpan = 1000.0d / 12.0d;
            for (int i = 0; i < 12; i++)
            {
                Path p = new Path();
                p.Data = Geometry.Parse("M 0,0 L -10,0 L -10,50 L 0,60 L 10,50 L 10,0 Z");//路径标记语法
                p.Opacity = 0.2;
                p.Fill = new SolidColorBrush(Colors.LightBlue);
                p.Stroke = new SolidColorBrush(Colors.DarkBlue);
                p.StrokeThickness = 1;

                TransformGroup tg = new TransformGroup();
                p.RenderTransform = tg;

                TranslateTransform tt = new TranslateTransform();//平移变换
                tt.Y = 50;

                RotateTransform rt = new RotateTransform();//旋转变换
                rt.Angle = i * 30;

                tg.Children.Add(tt);
                tg.Children.Add(rt);

                DoubleAnimation da = new DoubleAnimation();
                da.Duration = new Duration(TimeSpan.FromSeconds(1));
                da.From = 1.0;
                da.To = 0.2;
                da.BeginTime = TimeSpan.FromMilliseconds(dAnimationTimeSpan * i);
                da.RepeatBehavior = RepeatBehavior.Forever;
                g.Children.Add(p);
                p.BeginAnimation(Path.OpacityProperty, da);
            }
        }

        protected void OnIsBusyPropertyChanged(object oldValue, object newValue)
        {
            this.IsBusy = (bool)newValue;
        }
    }
}
