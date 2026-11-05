using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using WpfBinding = System.Windows.Data.Binding;

namespace UniclassClassifier.Views
{
    public static class ClassificationStyles
    {
        public static readonly System.Windows.Media.Color Background = System.Windows.Media.Color.FromRgb(247, 248, 250);
        public static readonly System.Windows.Media.Color Header = System.Windows.Media.Color.FromRgb(44, 62, 80);
        public static readonly System.Windows.Media.Color Level2 = System.Windows.Media.Color.FromRgb(142, 68, 173);
        public static readonly System.Windows.Media.Color Level1 = System.Windows.Media.Color.FromRgb(46, 134, 193);
        public static readonly System.Windows.Media.Color Review = System.Windows.Media.Color.FromRgb(230, 126, 34);
        public static readonly System.Windows.Media.Color CategoryCheck = System.Windows.Media.Color.FromRgb(192, 57, 43);
        public static readonly System.Windows.Media.Color Ok = System.Windows.Media.Color.FromRgb(39, 174, 96);
        public static readonly System.Windows.Media.Color Neutral = System.Windows.Media.Color.FromRgb(108, 122, 133);

        public static SolidColorBrush Brush(System.Windows.Media.Color color) => new SolidColorBrush(color);

        public static System.Windows.Media.Color StatusColor(string status) => status switch
        {
            "Review: category check" => CategoryCheck,
            "Review: low confidence" => Review,
            "Mixed codes" => Level1,
            _ => Ok
        };

        public static Button Button(string text, System.Windows.Media.Color color, RoutedEventHandler click)
        {
            var button = new Button
            {
                Content = text,
                Height = 40,
                MinWidth = 128,
                Padding = new Thickness(18, 0, 18, 0),
                Margin = new Thickness(8, 0, 0, 0),
                Background = Brush(color),
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Cursor = System.Windows.Input.Cursors.Hand,
                Template = ButtonTemplate()
            };
            button.Click += click;
            return button;
        }

        public static Border Card(UIElement content, string? title = null)
        {
            var panel = new StackPanel();
            if (!string.IsNullOrWhiteSpace(title))
                panel.Children.Add(new TextBlock { Text = title, FontSize = 16, FontWeight = FontWeights.Bold, Foreground = Brush(Header), Margin = new Thickness(0, 0, 0, 12) });
            panel.Children.Add(content);
            return new Border
            {
                Child = panel,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 230, 235)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(18),
                Margin = new Thickness(0, 0, 14, 14),
                Effect = new System.Windows.Media.Effects.DropShadowEffect { BlurRadius = 10, ShadowDepth = 1, Opacity = 0.08 }
            };
        }

        private static ControlTemplate ButtonTemplate()
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(5));
            border.SetBinding(Border.BackgroundProperty, new WpfBinding("Background") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            border.SetBinding(Border.PaddingProperty, new WpfBinding("Padding") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            content.SetValue(ContentPresenter.ContentSourceProperty, "Content");
            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }

        public sealed class StatusBrushConverter : IValueConverter
        {
            public object Convert(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                var status = value as string ?? "";
                return Brush(StatusColor(status));
            }

            public object ConvertBack(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture) => System.Windows.Data.Binding.DoNothing;
        }

        public sealed class StatusTextConverter : IValueConverter
        {
            public object Convert(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture) => value ?? "";
            public object ConvertBack(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture) => System.Windows.Data.Binding.DoNothing;
        }
    }
}
