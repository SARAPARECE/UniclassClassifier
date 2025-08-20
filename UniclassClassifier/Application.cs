using System.Reflection;
using Autodesk.Revit.UI;
using Nice3point.Revit.Toolkit.External;
using UniclassClassifier.Commands;

namespace UniclassClassifier
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            CreateRibbon();
        }

        private void CreateRibbon()
        {
            // Painel 1 – Classifier
            var panelClassifier = Application.CreatePanel("Classifier", "UniclassClassifier");

            var classifierButton = panelClassifier.AddPushButton<ClassifierCommand>("AI Classifier")
                .SetImage("/UniclassClassifier;component/Resources/Icons/icon.png")
                .SetLargeImage("/UniclassClassifier;component/Resources/Icons/icon.png");

            classifierButton.ToolTip = "Classify BIM elements using AI.";
            classifierButton.LongDescription = "Uses trained machine learning models to automatically classify BIM elements using Uniclass codes based on geometry and metadata.";

            // Painel 2 – Data
            var panelData = Application.CreatePanel("Data", "UniclassClassifier");

            // Caminho da DLL atual
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // Criar os botões Export e Info
            var exportBtnData = new PushButtonData("Export", "Export", assemblyPath, typeof(ExportCommand).FullName);
            var infoBtnData = new PushButtonData("Info", "Info", assemblyPath, typeof(InfoCommand).FullName);

            // Adicionar botões empilhados no painel Data
            var stackedButtons = panelData.AddStackedItems(exportBtnData, infoBtnData);

            // Botão Export
            var exportButton = (PushButton)stackedButtons[0];
            exportButton.ToolTip = "Export BIM element data.";
            exportButton.LongDescription = "Exports element geometry, classification and quantities to external files (CSV or JSON).";
            exportButton.SetImage("/UniclassClassifier;component/Resources/Icons/icon2.png");

            // Botão Info
            var infoButton = (PushButton)stackedButtons[1];
            infoButton.ToolTip = "Plugin information.";
            infoButton.LongDescription = "Displays plugin details including version, authorship and reference documentation.";
            infoButton.SetImage("/UniclassClassifier;component/Resources/Icons/icon3.png");
        }
    }
}
