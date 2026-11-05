using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CsvHelper;
using CsvHelper.Configuration;
using Newtonsoft.Json;
using UniclassClassifier.Views;

namespace UniclassClassifier.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ClassifierCommand : IExternalCommand
    {
        // ElementId as text (IntegerValue was replaced by Value in Revit 2024)
        private static string IdString(ElementId id)
        {
#if REVIT2024_OR_GREATER
            return id.Value.ToString();
#else
            return id.IntegerValue.ToString();
#endif
        }

        private const double DefaultTau = 0.8;

        // English category names as used in the training data, independent of the language of Revit
        private static readonly Dictionary<BuiltInCategory, string> CategoryNames = new Dictionary<BuiltInCategory, string>
        {
            [BuiltInCategory.OST_Walls] = "Walls",
            [BuiltInCategory.OST_Floors] = "Floors",
            [BuiltInCategory.OST_Roofs] = "Roofs",
            [BuiltInCategory.OST_StructuralColumns] = "Structural Columns",
            [BuiltInCategory.OST_StructuralFraming] = "Structural Framing",
            [BuiltInCategory.OST_Columns] = "Columns",
            [BuiltInCategory.OST_Doors] = "Doors",
            [BuiltInCategory.OST_Windows] = "Windows",
            [BuiltInCategory.OST_CurtainWallPanels] = "Curtain Panels",
            [BuiltInCategory.OST_Stairs] = "Stairs",
            [BuiltInCategory.OST_StructuralFoundation] = "Structural Foundations"
        };

        public static string CategoryName(BuiltInCategory bic, Element elem)
        {
            if (CategoryNames.TryGetValue(bic, out var name)) return name;
            var n = elem.Category?.Name;
            return string.IsNullOrWhiteSpace(n) ? "None" : n;
        }

        private class ClassificationSummary
        {
            public int n { get; set; }
            public int L2 { get; set; }
            public int L1 { get; set; }
            public int review { get; set; }
            public double tau_L2 { get; set; }
            public double tau_L1 { get; set; }
            public List<string>? review_ids { get; set; }
            public int category_check { get; set; }
        }

        // ElementId from the id written by the Python script (Int64 ids from Revit 2024)
        private static ElementId ToElementId(long v)
        {
#if REVIT2024_OR_GREATER
            return new ElementId(v);
#else
            return new ElementId((int)v);
#endif
        }
        private const double FeetToMeter = 0.3048;
        private const double Ft2ToM2 = 0.092903;
        private const double Ft3ToM3 = 0.0283168;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var exportTime = TimeSpan.Zero;
            var filePreparationTime = TimeSpan.Zero;
            var pythonTime = TimeSpan.Zero;
            var reportReadingTime = TimeSpan.Zero;
            var groupingTime = TimeSpan.Zero;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Confidence thresholds for Level 2 (Sub-group) and Level 1 (Group)
            var dialog = new ThresholdDialog(DefaultTau, DefaultTau);
            new System.Windows.Interop.WindowInteropHelper(dialog).Owner = commandData.Application.MainWindowHandle;
            if (dialog.ShowDialog() != true)
                return Result.Cancelled;

            try
            {
                string dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string tempDir = Path.Combine(dllDir, "Temp");
                string scriptPath = Path.Combine(dllDir, "Scripts", "predict_v1.py");

                Directory.CreateDirectory(tempDir);
                string jsonPath = Path.Combine(tempDir, "export.json");
                string csvPath = Path.Combine(tempDir, "export.csv");
                string classifiedPath = Path.Combine(tempDir, "classified.json");
                string modelDir = Path.Combine(dllDir, "Scripts", "model_v1");
                string summaryPath = Path.Combine(tempDir, "summary.json");
                string reportPath = Path.Combine(tempDir, "classified_report.csv");
                foreach (var old in new[] { classifiedPath, summaryPath, reportPath })
                    if (File.Exists(old)) File.Delete(old);

                var stage = Stopwatch.StartNew();
                var exportData = ExportElementData(doc);
                stage.Stop();
                exportTime = stage.Elapsed;

                stage.Restart();
                File.WriteAllText(jsonPath, JsonConvert.SerializeObject(exportData, Formatting.Indented));

                using (var writer = new StreamWriter(csvPath, false, new UTF8Encoding(true)))
                {
                    if (exportData.Count > 0)
                    {
                        var headers = exportData[0].Keys.ToList();
                        writer.WriteLine(string.Join(",", headers));

                        foreach (var row in exportData)
                        {
                            var values = headers.Select(h => SanitizeCsv(row.ContainsKey(h) ? row[h] : "None"));
                            writer.WriteLine(string.Join(",", values));
                        }
                    }
                }
                stage.Stop();
                filePreparationTime = stage.Elapsed;
                // Python environment with pandas, scikit-learn and xgboost (see Scripts/requirements_v1.txt);
                // the environment variable UNICLASS_PYTHON overrides the default path
                string pythonExe = Environment.GetEnvironmentVariable("UNICLASS_PYTHON")
                                   ?? @"C:\\Users\\Admin\\anaconda3\\envs\\Pyrevit\\python.exe";
                if (!File.Exists(pythonExe))
                    pythonExe = "python";   // Python on the PATH
                var psi = new ProcessStartInfo
                {
                    FileName = pythonExe,
                    Arguments = $"\"{scriptPath}\" \"{csvPath}\" \"{jsonPath}\" \"{modelDir}\" " +
                                $"{dialog.TauLevel2.ToString(CultureInfo.InvariantCulture)} {dialog.TauLevel1.ToString(CultureInfo.InvariantCulture)}",
                    UseShellExecute = false,  // ✅ TEM de estar aqui antes do redirecionamento
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                var pythonStopwatch = Stopwatch.StartNew();
                using (var process = Process.Start(psi))
                {
                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        TaskDialog.Show("Erro no script", string.IsNullOrWhiteSpace(stderr) ? stdout : stderr);
                        return Result.Failed;
                    }
                }
                pythonStopwatch.Stop();
                pythonTime = pythonStopwatch.Elapsed;

                if (!File.Exists(classifiedPath))
                {
                    TaskDialog.Show("Erro", "Ficheiro 'classified.json' não encontrado.");
                    return Result.Failed;
                }

                if (!File.Exists(reportPath))
                {
                    TaskDialog.Show("Erro", "Ficheiro 'classified_report.csv' não encontrado.");
                    return Result.Failed;
                }

                var titlesPath = Path.Combine(dllDir, "Scripts", "ss_titles.json");
                var titles = ReadTitles(titlesPath);
                stage.Restart();
                var reportRows = ReadReport(reportPath);
                stage.Stop();
                reportReadingTime = stage.Elapsed;
                stage.Restart();
                var typeRows = ClassificationGrouping.Build(reportRows, titles);
                stage.Stop();
                groupingTime = stage.Elapsed;
                var resultsWindow = new ClassificationWindow(uidoc, typeRows, titles, dialog.TauLevel2, dialog.TauLevel1, pythonTime,
                    new ClassificationTimings
                    {
                        Export = exportTime,
                        FilePreparation = filePreparationTime,
                        Python = pythonTime,
                        ReportReading = reportReadingTime,
                        Grouping = groupingTime
                    });
                new System.Windows.Interop.WindowInteropHelper(resultsWindow).Owner = commandData.Application.MainWindowHandle;
                resultsWindow.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                string dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string tempDir = Path.Combine(dllDir, "Temp");
                string scriptPath = Path.Combine(dllDir, "Scripts", "predict_v1.py");

                string detalhesErro = $"Erro: {ex.Message}\n\n" +
                                      $"Script: {scriptPath}\n" +
                                      $"Temp: {tempDir}\n\n" +
                                      $"Stack Trace:\n{ex.StackTrace}";

                TaskDialog.Show("Erro no Classifier", detalhesErro);
                return Result.Failed;
            }
        }

        private static List<ClassificationReportRow> ReadReport(string path)
        {
            var result = new List<ClassificationReportRow>();
            using var reader = new StreamReader(path, new UTF8Encoding(true));
            using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                BadDataFound = null,
                MissingFieldFound = null
            });
            csv.Read();
            csv.ReadHeader();
            while (csv.Read())
            {
                if (!long.TryParse(csv.GetField("ElementID"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var elementId))
                    continue;
                result.Add(new ClassificationReportRow
                {
                    ElementId = elementId,
                    FamilyAndType = csv.GetField("Family and Type") ?? "",
                    Category = csv.GetField("Category") ?? "",
                    Code = csv.GetField("code") ?? "",
                    Level = csv.GetField("level") ?? "",
                    L2Pred = csv.GetField("L2_pred") ?? "",
                    PL2 = ParseNumber(csv.GetField("p_L2")),
                    L1Pred = csv.GetField("L1_pred") ?? "",
                    PL1 = ParseNumber(csv.GetField("p_L1")),
                    CategoryGroup = csv.GetField("category_group") ?? "",
                    CategoryMismatch = bool.TryParse(csv.GetField("category_mismatch"), out var mismatch) && mismatch
                });
            }
            return result;
        }

        private static double ParseNumber(string? value) =>
            double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : 0;

        private static IReadOnlyDictionary<string, string> ReadTitles(string path)
        {
            if (!File.Exists(path)) return new Dictionary<string, string>();
            return JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(path, Encoding.UTF8))
                   ?? new Dictionary<string, string>();
        }

        private List<Dictionary<string, string>> ExportElementData(Document doc)
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Columns,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_CurtainWallPanels,
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_StructuralFoundation
            };

            var exportData = new List<Dictionary<string, string>>();

            foreach (var bic in categories)
            {
                var collector = new FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var elem in collector)
                {
                    var type = doc.GetElement(elem.GetTypeId()) as ElementType;
                    if (type == null) continue;

                    var bbox = elem.get_BoundingBox(null);
                    var coords = GetElementCoordinates(elem);
                    var centroid = GetCentroid(elem);
                    

                    var data = new Dictionary<string, string>
                    {
                        ["ElementID"] = IdString(elem.Id),
                        ["Family and Type"] = FormatString($"{type.FamilyName} - {type.Name}"),
                        ["Category"] = CategoryName(bic, elem),
                        ["Volume"] = GetDouble(elem, BuiltInParameter.HOST_VOLUME_COMPUTED, Ft3ToM3),
                        ["Area"] = GetDouble(elem, BuiltInParameter.HOST_AREA_COMPUTED, Ft2ToM2),
                        ["load_bearing_status"] = FormatString(GetStructuralStatus(elem)),
                        ["Length"] = GetDouble(elem, BuiltInParameter.CURVE_ELEM_LENGTH, FeetToMeter),
                        ["Height"] = TryGetHeight(elem, type),
                        ["Thickness/Width"] = TryGetWidthOrThickness(elem, type),
                        ["Total_Surface_Area"] = bbox != null ? FormatDouble((bbox.Max.X - bbox.Min.X) * (bbox.Max.Y - bbox.Min.Y) * Ft2ToM2) : "None",
                        ["Base_constraint"] = GetParamValue(elem, "Base Constraint"),
                        ["Top_constraint"] = GetParamValue(elem, "Top Constraint"),
                        ["Base_offset"] = GetParamValue(elem, "Base Offset"),
                        ["Top_offset"] = GetParamValue(elem, "Top Offset"),
                        ["Number_of_Faces"] = GetNumberOfFaces(elem),
                        ["Level"] = GetParamValue(elem, "Level"),
                        ["Phase_Created"] = GetParamValue(elem, "Phase Created"),
                        ["Start_X"] = FormatNullable(coords.startX),
                        ["Start_Y"] = FormatNullable(coords.startY),
                        ["Start_Z"] = FormatNullable(coords.startZ),
                        ["End_X"] = FormatNullable(coords.endX),
                        ["End_Y"] = FormatNullable(coords.endY),
                        ["End_Z"] = FormatNullable(coords.endZ),
                        ["Orientation_Angle"] = GetOrientationAngle(elem),
                        ["Curvature"] = GetCurvature(elem),
                        ["Bounding_Box_Width"] = bbox != null ? FormatDouble((bbox.Max.X - bbox.Min.X) * FeetToMeter) : "None",
                        ["Bounding_Box_Height"] = bbox != null ? FormatDouble((bbox.Max.Z - bbox.Min.Z) * FeetToMeter) : "None",
                        ["Bounding_Box_Depth"] = bbox != null ? FormatDouble((bbox.Max.Y - bbox.Min.Y) * FeetToMeter) : "None",
                        ["Centroid_X"] = centroid.x,
                        ["Centroid_Y"] = centroid.y,
                        ["Centroid_Z"] = centroid.z,
                        ["Materials"] = FormatString(string.Join(" | ", elem.GetMaterialIds(false).Select(id => doc.GetElement(id)?.Name).Where(n => !string.IsNullOrEmpty(n))))
                    };

                    exportData.Add(data);
                }
            }

            return exportData;
        }

        private (string x, string y, string z) GetCentroid(Element elem)
        {
            try
            {
                Options options = new Options { ComputeReferences = false, IncludeNonVisibleObjects = false, DetailLevel = ViewDetailLevel.Fine };
                GeometryElement geomElement = elem.get_Geometry(options);

                if (geomElement == null) return ("None", "None", "None");

                foreach (GeometryObject geomObj in geomElement)
                {
                    Solid solid = geomObj as Solid;
                    if (solid != null && solid.Volume > 0)
                    {
                        var centroid = solid.ComputeCentroid();
                        return (FormatDouble(centroid.X * FeetToMeter), FormatDouble(centroid.Y * FeetToMeter), FormatDouble(centroid.Z * FeetToMeter));
                    }
                }
            }
            catch { }
            return ("None", "None", "None");
        }

        private string GetDouble(Element elem, BuiltInParameter bip, double factor)
        {
            var p = elem.get_Parameter(bip);
            return (p != null && p.StorageType == StorageType.Double) ? FormatDouble(p.AsDouble() * factor) : "None";
        }

        private string GetParamValue(Element elem, string name) => FormatString(elem.LookupParameter(name)?.AsValueString());

        private string TryGetHeight(Element elem, ElementType type)
        {
            var names = new[] { "Height", "Unconnected Height", "Altura", "Head Height" };
            foreach (var n in names)
            {
                var p = elem.LookupParameter(n) ?? type.LookupParameter(n);
                if (p != null && p.StorageType == StorageType.Double)
                    return FormatDouble(p.AsDouble() * FeetToMeter);
            }
            return "None";
        }

        private string TryGetWidthOrThickness(Element elem, ElementType type)
        {
            var names = new[] { "Width", "Thickness" };
            foreach (var n in names)
            {
                var p = elem.LookupParameter(n) ?? type.LookupParameter(n);
                if (p != null && p.StorageType == StorageType.Double)
                    return FormatDouble(p.AsDouble() * FeetToMeter);
            }
            return "None";
        }

        private string GetStructuralStatus(Element elem)
        {
            try
            {
                var cat = elem.Category.Name;
                BuiltInParameter bip = cat switch
                {
                    "Walls" => BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT,
                    "Floors" => BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL,
                    _ => 0
                };
                var p = elem.get_Parameter(bip);
                return p?.AsInteger() == 1 ? "Load-Bearing" : "Non Load-Bearing";
            }
            catch { return "None"; }
        }

        private string GetNumberOfFaces(Element elem)
        {
            try
            {
                Options options = new Options { ComputeReferences = false, IncludeNonVisibleObjects = false, DetailLevel = ViewDetailLevel.Fine };
                GeometryElement geomElement = elem.get_Geometry(options);

                if (geomElement == null) return "None";

                int faceCount = 0;

                foreach (GeometryObject geomObj in geomElement)
                {
                    if (geomObj is Solid solid && solid.Faces != null)
                        faceCount += solid.Faces.Size;
                    else if (geomObj is GeometryInstance instance)
                    {
                        GeometryElement instGeom = instance.GetInstanceGeometry();
                        foreach (GeometryObject instObj in instGeom)
                        {
                            if (instObj is Solid instSolid && instSolid.Faces != null)
                                faceCount += instSolid.Faces.Size;
                        }
                    }
                }

                return faceCount > 0 ? faceCount.ToString() : "None";
            }
            catch
            {
                return "None";
            }
        }
        private string GetOrientationAngle(Element elem)
        {
            try
            {
                if (elem.Location is LocationCurve lc)
                {
                    XYZ start = lc.Curve.GetEndPoint(0);
                    XYZ end = lc.Curve.GetEndPoint(1);
                    XYZ direction = (end - start).Normalize();
                    double angleRad = Math.Atan2(direction.Y, direction.X); // ângulo no plano XY
                    double angleDeg = angleRad * (180.0 / Math.PI);
                    return FormatDouble(angleDeg);
                }
            }
            catch { }

            return "None";
        }
        private string GetCurvature(Element elem)
        {
            try
            {
                if (elem.Location is LocationCurve lc)
                {
                    Curve curve = lc.Curve;
                    return curve.IsCyclic ? "Cyclic" : curve.IsBound ? "Bound" : "Unbound";
                }
            }
            catch { }
            return "None";
        }
        private string FormatDouble(double value) => value.ToString(CultureInfo.InvariantCulture);
        private string FormatNullable(string value) => string.IsNullOrWhiteSpace(value) ? "None" : value;
        private string FormatString(string value) => string.IsNullOrWhiteSpace(value) ? "None" : value;
        private static string SanitizeCsv(string input) => $"\"{input?.Replace("\r", " ").Replace("\n", " ").Replace("\"", "\"\"") ?? "None"}\"";

        private (string startX, string startY, string startZ, string endX, string endY, string endZ) GetElementCoordinates(Element elem)
        {
            var loc = elem.Location;
            if (loc is LocationPoint lp)
            {
                var pt = lp.Point;
                return (FormatDouble(pt.X * FeetToMeter), FormatDouble(pt.Y * FeetToMeter), FormatDouble(pt.Z * FeetToMeter), "None", "None", "None");
            }
            if (loc is LocationCurve lc)
            {
                var start = lc.Curve.GetEndPoint(0);
                var end = lc.Curve.GetEndPoint(1);
                return (FormatDouble(start.X * FeetToMeter), FormatDouble(start.Y * FeetToMeter), FormatDouble(start.Z * FeetToMeter),
                        FormatDouble(end.X * FeetToMeter), FormatDouble(end.Y * FeetToMeter), FormatDouble(end.Z * FeetToMeter));
            }
            return ("None", "None", "None", "None", "None", "None");
        }
       
    }
}


