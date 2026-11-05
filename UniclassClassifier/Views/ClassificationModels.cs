using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System;

namespace UniclassClassifier.Views
{
    public sealed class ClassificationTimings
    {
        public TimeSpan Export { get; init; }
        public TimeSpan FilePreparation { get; init; }
        public TimeSpan Python { get; init; }
        public TimeSpan ReportReading { get; init; }
        public TimeSpan Grouping { get; init; }
        public TimeSpan WindowConstruction { get; set; }
    }

    public sealed class ClassificationReportRow
    {
        public long ElementId { get; set; }
        public string FamilyAndType { get; set; } = "";
        public string Category { get; set; } = "";
        public string Code { get; set; } = "";
        public string Level { get; set; } = "";
        public string L2Pred { get; set; } = "";
        public double PL2 { get; set; }
        public string L1Pred { get; set; } = "";
        public double PL1 { get; set; }
        public string CategoryGroup { get; set; } = "";
        public bool CategoryMismatch { get; set; }
    }

    public sealed class ClassificationTypeRow : INotifyPropertyChanged
    {
        private string _selectedCode = "";
        private bool _include;

        public string Key { get; init; } = "";
        public string Category { get; init; } = "";
        public string FamilyAndType { get; init; } = "";
        public List<long> ElementIds { get; init; } = new List<long>();
        public int InstanceCount => ElementIds.Count;
        public string Proposed { get; init; } = "";
        public double Confidence { get; init; }
        public string Status { get; init; } = "";
        public string Detail { get; init; } = "";
        public string MostLikelyLevel2 { get; init; } = "";
        public IReadOnlyList<ClassificationReportRow> SourceRows { get; init; } = Array.Empty<ClassificationReportRow>();
        public string Title(string code, IReadOnlyDictionary<string, string> titles) =>
            !string.IsNullOrWhiteSpace(code) && titles.TryGetValue(code, out var title) ? title : "";

        public string SelectedCode
        {
            get => _selectedCode;
            set
            {
                if (_selectedCode == value) return;
                _selectedCode = value ?? "";
                OnPropertyChanged();
                OnPropertyChanged(nameof(Description));
            }
        }

        public bool Include
        {
            get => _include;
            set
            {
                if (_include == value) return;
                _include = value;
                OnPropertyChanged();
            }
        }

        public IReadOnlyDictionary<string, string> Titles { get; init; } = new Dictionary<string, string>();
        public string Description => string.IsNullOrWhiteSpace(SelectedCode) ? "(not written)" : Title(SelectedCode, Titles);

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public static class ClassificationGrouping
    {
        public static List<ClassificationTypeRow> Build(IEnumerable<ClassificationReportRow> rows, IReadOnlyDictionary<string, string> titles)
        {
            return rows.GroupBy(r => r.Category + "\u001f" + r.FamilyAndType)
                .Select(g => BuildOne(g.Key, g.ToList(), titles))
                .OrderBy(r => StatusOrder(r.Status))
                .ThenByDescending(r => r.InstanceCount)
                .ToList();
        }

        private static ClassificationTypeRow BuildOne(string key, List<ClassificationReportRow> rows, IReadOnlyDictionary<string, string> titles)
        {
            var assigned = rows.Where(r => !string.Equals(r.Level, "L1 review", StringComparison.OrdinalIgnoreCase)).ToList();
            var codeCounts = assigned.Where(r => !string.IsNullOrWhiteSpace(r.Code))
                .GroupBy(r => r.Code).OrderByDescending(g => g.Count()).ThenBy(g => g.Key).ToList();
            var proposed = codeCounts.FirstOrDefault()?.Key ?? rows.GroupBy(r => r.L1Pred).OrderByDescending(g => g.Count()).FirstOrDefault()?.Key ?? "";
            var proposedRows = assigned.Where(r => r.Code == proposed).ToList();
            var isL2 = proposed.Length > 5;
            var confidenceRows = proposedRows.Count > 0 ? proposedRows : rows;
            var confidence = confidenceRows.Count == 0 ? 0 : confidenceRows.Average(r => isL2 ? r.PL2 : r.PL1);
            var mismatched = rows.FirstOrDefault(r => r.CategoryMismatch);
            var allReview = assigned.Count == 0;
            var hasReview = rows.Any(r => string.Equals(r.Level, "L1 review", StringComparison.OrdinalIgnoreCase));
            string status;
            if (mismatched != null) status = "Review: category check";
            else if (allReview) status = "Review: low confidence";
            else if (codeCounts.Count > 1 || hasReview) status = "Mixed codes";
            else status = "Automatic";

            var parts = codeCounts.Select(c => $"{c.Key} ({c.Count()})").ToList();
            var reviewCount = rows.Count(r => string.Equals(r.Level, "L1 review", StringComparison.OrdinalIgnoreCase));
            if (reviewCount > 0) parts.Add($"review ({reviewCount})");
            var detail = string.Join(", ", parts);
            if (mismatched != null)
                detail += $"; predicted {mismatched.L1Pred}, category usually {mismatched.CategoryGroup}";
            var likelyL2 = rows.GroupBy(r => r.L2Pred).OrderByDescending(g => g.Count()).FirstOrDefault()?.Key ?? "";

            var split = key.Split('\u001f');
            var row = new ClassificationTypeRow
            {
                Key = key,
                Category = split[0],
                FamilyAndType = split.Length > 1 ? split[1] : "",
                ElementIds = rows.Select(r => r.ElementId).ToList(),
                Proposed = proposed,
                Confidence = confidence,
                Status = status,
                Detail = detail + (string.IsNullOrWhiteSpace(likelyL2) ? "" : $"; most likely Level 2: {likelyL2}"),
                MostLikelyLevel2 = likelyL2,
                SourceRows = rows,
                Titles = titles,
                SelectedCode = status == "Automatic" || status == "Mixed codes" ? proposed : "",
            };
            row.Include = !string.IsNullOrWhiteSpace(row.SelectedCode);
            return row;
        }

        public static int StatusOrder(string status) => status switch
        {
            "Review: category check" => 0,
            "Review: low confidence" => 1,
            "Mixed codes" => 2,
            _ => 3
        };
    }
}
