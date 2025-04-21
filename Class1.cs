using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RevCloudInRed

{
    [Transaction(TransactionMode.Manual)]
    public class PrintWithColoredRevisionClouds : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)

        {
            System.Diagnostics.Debug.WriteLine("PrintWithColoredRevisionClouds");
            System.Diagnostics.Debugger.Launch();

            TaskDialog.Show("Debug", "Star of Execute");

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Collect all sheets
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet));

            List<ViewSheet> sheetsCollector = collector
                .Cast<ViewSheet>()
                .Where(sheet => !sheet.IsPlaceholder)
                .ToList();


            // Override Revision Cloud color in each sheet view
            { }
            Category revisionCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_RevisionClouds);

            if (revisionCategory == null)
            {
                TaskDialog.Show("Error", "Revision Cloud category not found.");
                return Result.Failed;
            }

            else if (revisionCategory != null)
            { 
                TaskDialog.Show("Debug", "Revision Cloud category found.");
            }

                using (Transaction tx = new Transaction(doc, "Override Revision Cloud Color"))
                {
                    tx.Start();
                    foreach (ViewSheet sheet in sheetsCollector)
                    {
                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        ogs.SetProjectionLineColor(new Color(255, 0, 0)); // Red
                        sheet.SetCategoryOverrides(revisionCategory.Id, ogs);

                    // Set the color for the revision clouds in the view
                    // OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                    //ogs.SetProjectionLineColor(new Color(255, 0, 0)); // Red
                    //sheet.SetCategoryOverrides(revisionCategory.Id, ogs);

                }
                tx.Commit();
                }

            // Setup print settings
            PrintManager printManager = doc.PrintManager;
            printManager.SelectNewPrintDriver("Microsoft Print to PDF");
            printManager.PrintRange = PrintRange.Select;
            printManager.PrintToFile = true;

            using (Transaction printTx = new Transaction(doc, "Configure Print Settings"))
            {
                printTx.Start();
                PrintSetup setup = printManager.PrintSetup;
                PrintSetting printSetting = setup.CurrentPrintSetting as PrintSetting;
                PrintParameters parameters = printSetting.PrintParameters;

                parameters.ColorDepth = ColorDepthType.Color;
                parameters.PaperPlacement = PaperPlacementType.Center;
                parameters.ZoomType = ZoomType.FitToPage;
                parameters.HideCropBoundaries = true;

                printTx.Commit();
            }

            // Choose a directory to save the PDFs
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string outputFolder = Path.Combine("C:\\Users\\rajith.r\\OneDrive - Setty and Associates\\Desktop", "Revit Sheet PDFs");
            Directory.CreateDirectory(outputFolder);

            // Print each sheet to PDF with custom filename
            foreach (ViewSheet sheet in sheetsCollector)
            {
                ViewSet vs = new ViewSet();
                vs.Insert(sheet);

                printManager.ViewSheetSetting.CurrentViewSheetSet.Views = vs;
                printManager.Apply();

                // Create a custom file name using Sheet Number and Name
                string fileName = $"{sheet.SheetNumber}_{sheet.Name}.pdf";
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    fileName = fileName.Replace(c, '_');
                }
                string filePath = Path.Combine(outputFolder, fileName);

                printManager.PrintToFileName = filePath;

                try
                {
                    printManager.SubmitPrint();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Print Error", $"Failed to print {sheet.Name}: {ex.Message}");
                }
                TaskDialog.Show("Check", $"Applied color override to sheet: {sheet.Name}");

            }

            TaskDialog.Show("Success", $"All sheets printed");

            //TaskDialog.Show("Success", $"All sheets printed to:\n{outputFolder}");
            return Result.Succeeded;
        }
    }
}
