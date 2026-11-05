using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace UniclassClassifier.Views
{
    /// <summary>
    ///     Asks for the confidence thresholds of the Uniclass Ss classification before it runs.
    /// </summary>
    public class ThresholdDialog : Window
    {
        private readonly TextBox _level2;
        private readonly TextBox _level1;

        public double TauLevel2 { get; private set; }
        public double TauLevel1 { get; private set; }

        public ThresholdDialog(double tauLevel2, double tauLevel1)
        {
            Title = "Uniclass Classifier";
            SizeToContent = SizeToContent.WidthAndHeight;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            var panel = new StackPanel { Margin = new Thickness(16) };
            panel.Children.Add(new TextBlock
            {
                Text = "Confidence thresholds (0 to 1)",
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 8)
            });
            _level2 = AddRow(panel, "Sub-group (Level 2):", tauLevel2);
            _level1 = AddRow(panel, "Group (Level 1):", tauLevel1);
            panel.Children.Add(new TextBlock
            {
                Text = "A Level 2 code is written when its confidence reaches the Level 2 threshold, otherwise the Level 1 code when it reaches the Level 1 threshold. Elements below both thresholds keep their current code and are selected for review.",
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 340,
                Margin = new Thickness(0, 8, 0, 12)
            });

            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = "Classify", Width = 80, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            ok.Click += OnOk;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            panel.Children.Add(buttons);

            Content = panel;
        }

        private static TextBox AddRow(System.Windows.Controls.Panel parent, string label, double value)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 4, 0, 4) };
            row.Children.Add(new TextBlock { Text = label, Width = 160, VerticalAlignment = VerticalAlignment.Center });
            var box = new TextBox { Text = value.ToString("0.00", CultureInfo.InvariantCulture), Width = 80 };
            row.Children.Add(box);
            parent.Children.Add(row);
            return box;
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            if (!TryParse(_level2.Text, out var t2) || !TryParse(_level1.Text, out var t1))
            {
                MessageBox.Show("Enter values between 0 and 1.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TauLevel2 = t2;
            TauLevel1 = t1;
            DialogResult = true;
        }

        private static bool TryParse(string text, out double value)
        {
            return double.TryParse(text.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                   && value >= 0 && value <= 1;
        }
    }
}
