using System.Reflection;
using Autodesk.Revit.UI;
using Nice3point.Revit.Toolkit.External;
using UniclassClassifier.Commands;
using System.Windows.Media.Imaging;

namespace UniclassClassifier
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            CreateRibbon();
        }

        private void CreateRibbon()
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            Application.CreateRibbonTab("Uniclass v2");

            // Painel 1 – Classifier
            var panelClassifier = Application.CreateRibbonPanel("Uniclass v2", "Classifier v2");

            var classifierButtonData = new PushButtonData(
                "ClassifierV2",
                "AI Classifier v2",
                assemblyPath,
                typeof(ClassifierCommand).FullName);
            var classifierButton = (PushButton)panelClassifier.AddItem(classifierButtonData);
            var classifierImage = new BitmapImage(new Uri("pack://application:,,,/UniclassClassifier_v2;component/Resources/Icons/icon.png"));
            classifierButton.Image = classifierImage;
            classifierButton.LargeImage = classifierImage;

            classifierButton.ToolTip = "Classify BIM elements using AI.";
            classifierButton.LongDescription = "Version 2: assigns Uniclass Ss Level 1 and Level 2 codes with a model that does not use the Revit category, a confidence gate and a check of the stated category.";

            // Painel 2 – Data
            var panelData = Application.CreateRibbonPanel("Uniclass v2", "Data");

            // Criar os botões Export e Info
            var exportBtnData = new PushButtonData("ExportV2", "Export", assemblyPath, typeof(ExportCommand).FullName);
            var infoBtnData = new PushButtonData("InfoV2", "Info", assemblyPath, typeof(InfoCommand).FullName);

            // Adicionar botões empilhados no painel Data
            var stackedButtons = panelData.AddStackedItems(exportBtnData, infoBtnData);

            // Botão Export
            var exportButton = (PushButton)stackedButtons[0];
            exportButton.ToolTip = "Export BIM element data.";
            exportButton.LongDescription = "Exports element geometry, classification and quantities to external files (CSV or JSON).";
            exportButton.Image = new BitmapImage(new Uri("pack://application:,,,/UniclassClassifier_v2;component/Resources/Icons/icon2.png"));

            // Botão Info
            var infoButton = (PushButton)stackedButtons[1];
            infoButton.ToolTip = "Plugin information.";
            infoButton.LongDescription = "Displays plugin details including version, authorship and reference documentation.";
            infoButton.Image = new BitmapImage(new Uri("pack://application:,,,/UniclassClassifier_v2;component/Resources/Icons/icon3.png"));
        }
    }
}
