using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using WpfBinding = System.Windows.Data.Binding;
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfGrid = System.Windows.Controls.Grid;

namespace UniclassClassifier.Views
{
    public sealed class ClassificationWindow : Window
    {
        private readonly UIDocument _uidoc;
        private readonly IReadOnlyDictionary<string, string> _titles;
        private readonly double _tauL2;
        private readonly double _tauL1;
        private readonly TimeSpan _pythonTime;
        private readonly ClassificationTimings _timings;
        private readonly ObservableCollection<ClassificationTypeRow> _rows;
        private readonly StackPanel _content = new StackPanel();
        private readonly TextBlock _stepText = new TextBlock();
        private readonly TextBlock _footer = new TextBlock();
        private Button? _applyButton;
        private int _step;

        public ClassificationWindow(UIDocument uidoc, IEnumerable<ClassificationTypeRow> rows, IReadOnlyDictionary<string, string> titles, double tauL2, double tauL1, TimeSpan pythonTime, ClassificationTimings timings)
        {
            _uidoc = uidoc;
            _titles = titles;
            _tauL2 = tauL2;
            _tauL1 = tauL1;
            _pythonTime = pythonTime;
            _timings = timings;
            _rows = new ObservableCollection<ClassificationTypeRow>(rows);
            foreach (var row in _rows)
                row.PropertyChanged += (_, __) => UpdateFooter();
            Title = "Uniclass Classifier v2";
            Width = 1180;
            Height = 760;
            MinWidth = 1000;
            MinHeight = 700;
            FontFamily = new FontFamily("Segoe UI");
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = ClassificationStyles.Brush(ClassificationStyles.Background);
            var construction = System.Diagnostics.Stopwatch.StartNew();
            BuildShell();
            ShowStep(0);
            construction.Stop();
            _timings.WindowConstruction = construction.Elapsed;
            System.Diagnostics.Debug.WriteLine(
                $"Uniclass startup: export={_timings.Export.TotalSeconds:0.000}s, files={_timings.FilePreparation.TotalSeconds:0.000}s, " +
                $"python={_timings.Python.TotalSeconds:0.000}s, report={_timings.ReportReading.TotalSeconds:0.000}s, " +
                $"grouping={_timings.Grouping.TotalSeconds:0.000}s, wpf={_timings.WindowConstruction.TotalSeconds:0.000}s");
        }

        private void BuildShell()
        {
            var root = new WpfGrid { Margin = new Thickness(26, 22, 26, 20) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var header = new StackPanel();
            var title = new TextBlock { Text = "Uniclass v2 classification", FontSize = 26, FontWeight = FontWeights.Bold, Foreground = ClassificationStyles.Brush(ClassificationStyles.Header), Margin = new Thickness(0, 0, 0, 16) };
            header.Children.Add(title);
            _stepText.FontSize = 13;
            _stepText.Margin = new Thickness(0, 0, 0, 22);
            header.Children.Add(_stepText);
            WpfGrid.SetRow(header, 0);
            root.Children.Add(header);
            _content.Margin = new Thickness(0, 0, 0, 12);
            _content.HorizontalAlignment = HorizontalAlignment.Stretch;
            WpfGrid.SetRow(_content, 1);
            root.Children.Add(_content);
            var footer = new DockPanel { LastChildFill = true, Margin = new Thickness(0, 8, 0, 0) };
            _footer.VerticalAlignment = VerticalAlignment.Center;
            _footer.FontSize = 12;
            _footer.Foreground = ClassificationStyles.Brush(ClassificationStyles.Header);
            footer.Children.Add(_footer);
            WpfGrid.SetRow(footer, 2);
            root.Children.Add(footer);
            Content = root;
        }

        private void ShowStep(int step)
        {
            _step = step;
            _content.Children.Clear();
            _stepText.Text = "1. Summary     2. Review     3. Confirm";
            _stepText.Inlines.Clear();
            for (var i = 0; i < 3; i++)
            {
                var run = new System.Windows.Documents.Run($"{i + 1}. {(i == 0 ? "Summary" : i == 1 ? "Review" : "Confirm")}") { FontWeight = i == step ? FontWeights.Bold : FontWeights.Normal, Foreground = ClassificationStyles.Brush(i == step ? ClassificationStyles.Header : i < step ? ClassificationStyles.Ok : ClassificationStyles.Neutral) };
                _stepText.Inlines.Add(run);
                if (i < 2) _stepText.Inlines.Add(new System.Windows.Documents.Run("    "));
            }
            if (step == 0) BuildSummary();
            else if (step == 1) BuildReview();
            else BuildConfirm();
        }

        private void BuildSummary()
        {
            var grid = new System.Windows.Controls.Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var left = new StackPanel();
            var total = _rows.Sum(r => r.InstanceCount);
            var l2 = _rows.Sum(r => r.SourceRows.Count(x => x.Level == "L2"));
            var l1 = _rows.Sum(r => r.SourceRows.Count(x => x.Level == "L1"));
            var review = _rows.Sum(r => r.SourceRows.Count(x => x.Level == "L1 review"));
            left.Children.Add(ClassificationStyles.Card(new TextBlock { Text = $"Elements processed: {total}\nFamily/types: {_rows.Count}\nProcessing time: {_pythonTime.TotalSeconds:0.00}s", TextWrapping = TextWrapping.Wrap }, "Overview"));
            var results = new StackPanel();
            results.Children.Add(new TextBlock { Text = "Level 2 code", Foreground = ClassificationStyles.Brush(ClassificationStyles.Level2) });
            results.Children.Add(Bar(l2, total, ClassificationStyles.Level2));
            results.Children.Add(new TextBlock { Text = "Level 1 code only", Foreground = ClassificationStyles.Brush(ClassificationStyles.Level1) });
            results.Children.Add(Bar(l1, total, ClassificationStyles.Level1));
            results.Children.Add(new TextBlock { Text = "Left for review", Foreground = ClassificationStyles.Brush(ClassificationStyles.Review) });
            results.Children.Add(Bar(review, total, ClassificationStyles.Review));
            var categoryMismatchElements = _rows.Sum(r => r.SourceRows.Count(x => x.CategoryMismatch));
            results.Children.Add(new TextBlock { Text = $"of which flagged by the category check: {categoryMismatchElements}", Foreground = ClassificationStyles.Brush(ClassificationStyles.CategoryCheck), Margin = new Thickness(0, 10, 0, 0) });
            left.Children.Add(ClassificationStyles.Card(results, "Classification results"));
            grid.Children.Add(left);
            var right = new StackPanel();
            right.Margin = new Thickness(8, 0, 0, 0);
            var groups = _rows.SelectMany(r => r.SourceRows).GroupBy(r => r.L1Pred).OrderBy(g => g.Key).Select(g => $"{g.Key}  {(_titles.TryGetValue(g.Key, out var t) ? t : "")}    L2: {g.Count(x => x.Level == "L2")}   L1: {g.Count(x => x.Level == "L1")}   Review: {g.Count(x => x.Level == "L1 review")}");
            right.Children.Add(ClassificationStyles.Card(new TextBlock { Text = string.Join("\n", groups), TextWrapping = TextWrapping.Wrap }, "Codes by system group"));
            right.Children.Add(ClassificationStyles.Card(new TextBlock { Text = $"Level 2 threshold: {_tauL2.ToString("0.00", CultureInfo.InvariantCulture)}\nLevel 1 threshold: {_tauL1.ToString("0.00", CultureInfo.InvariantCulture)}\nv1.1, no Revit category as input\nCategory check: On" }, "Configuration used"));
            System.Windows.Controls.Grid.SetColumn(right, 1);
            grid.Children.Add(right);
            _content.Children.Add(grid);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ClassificationStyles.Button("Close without writing", ClassificationStyles.Neutral, (_, __) => Close()));
            buttons.Children.Add(ClassificationStyles.Button("Select review elements", ClassificationStyles.Review, (_, __) => SelectRows(_rows.Where(r => r.Status != "Automatic"))));
            buttons.Children.Add(ClassificationStyles.Button("Next: review >", ClassificationStyles.Header, (_, __) => ShowStep(1)));
            _content.Children.Add(buttons);
            UpdateFooter();
        }

        private static Border Bar(int count, int total, System.Windows.Media.Color color)
        {
            var percent = total == 0 ? 0 : count * 100.0 / total;
            var bar = new Border { Height = 28, Background = ClassificationStyles.Brush(color), CornerRadius = new CornerRadius(4), Margin = new Thickness(0, 3, 0, 8), Child = new TextBlock { Text = $"{percent:0.0}% ({count})", Foreground = Brushes.White, FontWeight = FontWeights.Bold, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center } };
            return bar;
        }

        private void BuildReview()
        {
            _content.Children.Add(new TextBlock { Text = "Inconsistencies and items for review", FontSize = 18, FontWeight = FontWeights.Bold, Foreground = ClassificationStyles.Brush(ClassificationStyles.Header), Margin = new Thickness(0, 0, 0, 10) });
            var legend = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
            AddLegendItem(legend, "Review: category check", "category mismatch");
            AddLegendItem(legend, "Review: low confidence", "below confidence threshold");
            AddLegendItem(legend, "Mixed codes", "more than one proposed code");
            _content.Children.Add(legend);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 0, 10) };
            buttons.Children.Add(ClassificationStyles.Button("< Back", ClassificationStyles.Neutral, (_, __) => ShowStep(0)));
            buttons.Children.Add(ClassificationStyles.Button("Select all review elements", ClassificationStyles.Review, (_, __) => SelectRows(_rows.Where(r => r.Status != "Automatic"))));
            buttons.Children.Add(ClassificationStyles.Button("Next: confirm >", ClassificationStyles.Header, (_, __) => ShowStep(2)));
            _content.Children.Add(buttons);
            var grid = MakeGrid(false);
            grid.ItemsSource = _rows.Where(r => r.Status != "Automatic");
            _content.Children.Add(grid);
            UpdateFooter();
        }

        private void BuildConfirm()
        {
            _content.Children.Add(new TextBlock { Text = "Write to: Classification.Uniclass.Ss.Number (fallback: Comments)", Foreground = ClassificationStyles.Brush(ClassificationStyles.Header), Margin = new Thickness(0, 0, 0, 8) });
            var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            actions.Children.Add(ClassificationStyles.Button("Accept all proposals", ClassificationStyles.Ok, (_, __) => { foreach (var r in _rows) { r.SelectedCode = r.Proposed; r.Include = !string.IsNullOrWhiteSpace(r.SelectedCode); } UpdateFooter(); }));
            actions.Children.Add(ClassificationStyles.Button("Clear review rows", ClassificationStyles.Neutral, (_, __) => { foreach (var r in _rows.Where(r => r.Status != "Automatic")) { r.SelectedCode = ""; r.Include = false; } UpdateFooter(); }));
            _content.Children.Add(actions);
            var grid = MakeGrid(true);
            grid.ItemsSource = _rows;
            _content.Children.Add(grid);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ClassificationStyles.Button("< Back", ClassificationStyles.Neutral, (_, __) => ShowStep(1)));
            buttons.Children.Add(ClassificationStyles.Button("Cancel", ClassificationStyles.Neutral, (_, __) => Close()));
            _applyButton = ClassificationStyles.Button("Apply to 0 types", ClassificationStyles.Ok, Apply);
            buttons.Children.Add(_applyButton);
            _content.Children.Add(buttons);
            UpdateFooter();
        }

        private DataGrid MakeGrid(bool confirm)
        {
            var grid = new DataGrid { AutoGenerateColumns = false, IsReadOnly = !confirm, CanUserAddRows = false, Height = 430, HeadersVisibility = DataGridHeadersVisibility.Column, RowBackground = Brushes.White, AlternatingRowBackground = ClassificationStyles.Brush(ClassificationStyles.Background), Margin = new Thickness(0, 0, 0, 12), GridLinesVisibility = DataGridGridLinesVisibility.Horizontal, BorderBrush = ClassificationStyles.Brush(System.Windows.Media.Color.FromRgb(221, 225, 230)), BorderThickness = new Thickness(1), FontSize = 12, RowHeight = 42, ColumnHeaderHeight = 34 };
            ScrollViewer.SetHorizontalScrollBarVisibility(grid, ScrollBarVisibility.Disabled);
            grid.LoadingRow += (_, args) =>
            {
                if (args.Row.Item is ClassificationTypeRow row)
                    args.Row.Background = row.Status == "Review: category check"
                        ? ClassificationStyles.Brush(System.Windows.Media.Color.FromRgb(253, 242, 241))
                        : row.Status == "Review: low confidence"
                            ? ClassificationStyles.Brush(System.Windows.Media.Color.FromRgb(254, 248, 241))
                            : row.Status == "Mixed codes"
                                ? ClassificationStyles.Brush(System.Windows.Media.Color.FromRgb(246, 250, 253))
                                : Brushes.White;
            };
            if (confirm) grid.Columns.Add(new DataGridCheckBoxColumn { Header = "Include", Width = 62, Binding = new WpfBinding("Include") { Mode = BindingMode.TwoWay } });
            var statusTemplate = new DataTemplate();
            var statusBorder = new FrameworkElementFactory(typeof(Border));
            statusBorder.SetValue(Border.CornerRadiusProperty, new CornerRadius(10));
            statusBorder.SetValue(Border.PaddingProperty, new Thickness(8, 3, 8, 3));
            statusBorder.SetValue(Border.HorizontalAlignmentProperty, HorizontalAlignment.Left);
            statusBorder.SetBinding(Border.BackgroundProperty, new WpfBinding("Status") { Converter = new ClassificationStyles.StatusBrushConverter() });
            var statusText = new FrameworkElementFactory(typeof(TextBlock));
            statusText.SetValue(TextBlock.ForegroundProperty, Brushes.White);
            statusText.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
            statusText.SetBinding(TextBlock.TextProperty, new WpfBinding("Status"));
            statusBorder.AppendChild(statusText);
            statusTemplate.VisualTree = statusBorder;
            grid.Columns.Add(new DataGridTemplateColumn { Header = "Status", Width = confirm ? 150 : 185, CellTemplate = statusTemplate });
            grid.Columns.Add(new DataGridTextColumn { Header = "Category", Width = 90, Binding = new WpfBinding("Category") });
            grid.Columns.Add(new DataGridTextColumn { Header = "Family and Type", Width = confirm ? new DataGridLength(250) : new DataGridLength(1, DataGridLengthUnitType.Star), Binding = new WpfBinding("FamilyAndType"), ElementStyle = WrappedTextStyle() });
            grid.Columns.Add(new DataGridTextColumn { Header = "Inst.", Width = 52, Binding = new WpfBinding("InstanceCount") });
            grid.Columns.Add(new DataGridTextColumn { Header = "Proposed", Width = 105, Binding = new WpfBinding("Proposed") });
            grid.Columns.Add(new DataGridTextColumn { Header = "Conf.", Width = 64, Binding = new WpfBinding("Confidence") { StringFormat = "0.00" } });
            grid.Columns.Add(new DataGridTextColumn { Header = "Detail", Width = confirm ? new DataGridLength(1, DataGridLengthUnitType.Star) : new DataGridLength(1, DataGridLengthUnitType.Star), Binding = new WpfBinding("Detail"), ElementStyle = WrappedTextStyle() });
            if (confirm)
            {
                var template = new DataTemplate();
                var combo = new FrameworkElementFactory(typeof(WpfComboBox));
                combo.SetValue(WpfComboBox.IsEditableProperty, true);
                combo.SetValue(WpfComboBox.ItemsSourceProperty, _titles.Keys.OrderBy(x => x).ToList());
                combo.SetBinding(WpfComboBox.TextProperty, new WpfBinding("SelectedCode") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
                combo.AddHandler(WpfComboBox.SelectionChangedEvent, new SelectionChangedEventHandler((s, e) => { if (s is WpfComboBox c && c.DataContext is ClassificationTypeRow r) r.Include = !string.IsNullOrWhiteSpace(r.SelectedCode); }));
                template.VisualTree = combo;
                grid.Columns.Add(new DataGridTemplateColumn { Header = "Selected code", Width = 145, CellTemplate = template });
                grid.Columns.Add(new DataGridTextColumn { Header = "Description", Width = new DataGridLength(1, DataGridLengthUnitType.Star), Binding = new WpfBinding("Description"), ElementStyle = WrappedTextStyle() });
                grid.CellEditEnding += (_, __) => UpdateFooter();
            }
            else
            {
                var selectTemplate = new DataTemplate();
                var selectButton = new FrameworkElementFactory(typeof(Button));
                selectButton.SetValue(Button.ContentProperty, "Select in model");
                selectButton.SetValue(Button.HeightProperty, 28.0);
                selectButton.SetValue(Button.BackgroundProperty, ClassificationStyles.Brush(ClassificationStyles.Header));
                selectButton.SetValue(Button.ForegroundProperty, Brushes.White);
                selectButton.SetValue(Button.BorderThicknessProperty, new Thickness(0));
                selectButton.AddHandler(Button.ClickEvent, new RoutedEventHandler((s, e) =>
                {
                    if (s is Button b && b.DataContext is ClassificationTypeRow row) SelectRows(new[] { row });
                }));
                selectTemplate.VisualTree = selectButton;
                grid.Columns.Add(new DataGridTemplateColumn { Header = "Actions", Width = 120, CellTemplate = selectTemplate });
            }
            return grid;
        }

        private static Style WrappedTextStyle()
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
            style.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
            style.Setters.Add(new Setter(TextBlock.MarginProperty, new Thickness(4, 2, 4, 2)));
            return style;
        }

        private static void AddLegendItem(WrapPanel legend, string status, string explanation)
        {
            var item = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 18, 6) };
            item.Children.Add(new Border
            {
                Background = ClassificationStyles.Brush(ClassificationStyles.StatusColor(status)),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(9, 3, 9, 3),
                Child = new TextBlock { Text = status, Foreground = Brushes.White, FontWeight = FontWeights.Bold }
            });
            item.Children.Add(new TextBlock { Text = explanation, Foreground = ClassificationStyles.Brush(System.Windows.Media.Color.FromRgb(105, 115, 125)), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0, 0, 0) });
            legend.Children.Add(item);
        }

        private void SelectRows(IEnumerable<ClassificationTypeRow> rows)
        {
            var ids = rows.SelectMany(r => r.ElementIds).Distinct().Select(ToElementId).Where(id => id != ElementId.InvalidElementId && _uidoc.Document.GetElement(id) != null).ToList();
            _uidoc.Selection.SetElementIds(ids);
            _footer.Text = $"{ids.Count} elements selected in the model.";
        }

        private void UpdateFooter()
        {
            if (_step != 2)
            {
                _footer.Text = "";
                return;
            }
            var selected = _rows.Where(r => r.Include && !string.IsNullOrWhiteSpace(r.SelectedCode)).ToList();
            _footer.Text = $"{selected.Count} family/types selected ({selected.Sum(r => r.InstanceCount)} elements). Undo with Ctrl+Z in Revit (one transaction).";
            if (_applyButton != null)
                _applyButton.Content = $"Apply to {selected.Count} types";
        }

        private void Apply(object sender, RoutedEventArgs e)
        {
            var selected = _rows.Where(r => r.Include && !string.IsNullOrWhiteSpace(r.SelectedCode)).ToList();
            int types = 0, elements = 0, comments = 0, skipped = 0;
            using (var tx = new Transaction(_uidoc.Document, "Uniclass v2: apply codes"))
            {
                tx.Start();
                foreach (var row in selected)
                {
                    var typeIds = new HashSet<ElementId>();
                    foreach (var rawId in row.ElementIds)
                    {
                        var element = _uidoc.Document.GetElement(ToElementId(rawId));
                        if (element == null) { skipped++; continue; }
                        var type = _uidoc.Document.GetElement(element.GetTypeId());
                        if (type == null) { skipped++; continue; }
                        var typeParameter = type.LookupParameter("Classification.Uniclass.Ss.Number");
                        if (typeParameter != null && !typeParameter.IsReadOnly) { typeParameter.Set(row.SelectedCode); typeIds.Add(type.Id); }
                        var instanceParameter = element.LookupParameter("Classification.Uniclass.Ss.Number");
                        if (instanceParameter != null && !instanceParameter.IsReadOnly) { instanceParameter.Set(row.SelectedCode); elements++; }
                        else if (typeParameter == null || typeParameter.IsReadOnly)
                        {
                            var fallback = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                            if (fallback != null && !fallback.IsReadOnly) { fallback.Set(row.SelectedCode); comments++; } else skipped++;
                        }
                    }
                    types += typeIds.Count;
                }
                tx.Commit();
            }
            TaskDialog.Show("Uniclass Classifier v2", $"Types written: {types}\nElements written: {elements}\nElements written to Comments: {comments}\nElements skipped: {skipped}");
            DialogResult = true;
        }

        private static ElementId ToElementId(long id)
        {
#if REVIT2024_OR_GREATER
            return new ElementId(id);
#else
            return new ElementId((int)id);
#endif
        }
    }
}
