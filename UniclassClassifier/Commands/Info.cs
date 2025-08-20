using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms; // para SaveFileDialog
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
                TaskDialog.Show("Título da Janela", "Texto da mensagem principal");
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