using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using Newtonsoft.Json;

namespace UniclassClassifier.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class InfoCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                TaskDialog.Show("Uniclass Classifier v2",
                    "Uniclass Classifier v2.0 (model 1.1)\n\n" +
                    "Assigns Uniclass Ss codes (Level 1 Group and Level 2 Sub-group) to walls, floors, roofs, doors, windows, " +
                    "curtain panels, stairs and structural elements.\n\n" +
                    "Model: XGBoost on bounding-box geometry and the base quantities Length, Height, Thickness/Width and Area, " +
                    "averaged with a logistic regression on the family/type name, trained on seven BIM models (26 Level 2 classes " +
                    "after label review). The Revit category is not an input; all inputs have IFC equivalents.\n\n" +
                    "Confidence gate: Level 2 code if its confidence reaches the Level 2 threshold, otherwise Level 1 code if it " +
                    "reaches the Level 1 threshold; other elements are left for review. Elements whose predicted Level 1 group differs " +
                    "from the usual group of their Revit category are also left for review (category check).\n\n" +
                    "Ferreira, Parece, Resende and Elvas, ISCTE-IUL, ISTAR.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}