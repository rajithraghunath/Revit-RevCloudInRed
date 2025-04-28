using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace RevCloudInRed
{
    [Transaction(TransactionMode.Manual)]
    public class PrintWithColoredRevisionClouds : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get selected ViewSheets only
            List<ViewSheet> sheetsCollector = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewSheet))
                            .Cast<ViewSheet>()
                            .Where(sheet => !sheet.IsPlaceholder)
                            .ToList();

            if (!sheetsCollector.Any())
            {
                TaskDialog.Show("No Sheets Selected", "Please select one or more sheets before running the command.");
                return Result.Cancelled;
            }

            List<ElementId> tempFilterIds = new List<ElementId>();

            using (Transaction tx = new Transaction(doc, "Create Black Override Filters"))
            {
                tx.Start();

                foreach (ViewSheet sheet in sheetsCollector)
                {
                    List<ElementId> categoryIdsToOverride = new List<ElementId>();
                    foreach (Category cat in doc.Settings.Categories)
                    {
                        if ((cat.CategoryType == CategoryType.Annotation || cat.CategoryType == CategoryType.Model) &&
                            cat.Id.IntegerValue != (int)BuiltInCategory.OST_RevisionClouds)
                        {
                            categoryIdsToOverride.Add(cat.Id);
                        }
                    }

                    BuiltInCategory[] mustInclude = new[]
                    {
                        BuiltInCategory.OST_CutOutlines,
                        BuiltInCategory.OST_Doors,
                        BuiltInCategory.OST_Materials,
                        BuiltInCategory.OST_Rooms,
                        BuiltInCategory.OST_FillPatterns,
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_FilledRegion,
                        BuiltInCategory.OST_WallsCutPattern,
                        BuiltInCategory.OST_WallsDefault,
                        BuiltInCategory.OST_WallsFinish1,
                        BuiltInCategory.OST_WallsFinish2,
                        BuiltInCategory.OST_WallsInsulation,
                        BuiltInCategory.OST_WallsMembrane,
                        BuiltInCategory.OST_WallsProjectionOutlines,
                        BuiltInCategory.OST_WallsStructure,
                        BuiltInCategory.OST_WallsSubstrate,
                        BuiltInCategory.OST_WallsSurfacePattern,
                        BuiltInCategory.OST_StackedWalls,
                        BuiltInCategory.OST_Windows
                    };

                    foreach (BuiltInCategory bic in mustInclude)
                    {
                        Category extraCat = Category.GetCategory(doc, bic);
                        if (extraCat != null && !categoryIdsToOverride.Contains(extraCat.Id))
                        {
                            categoryIdsToOverride.Add(extraCat.Id);
                        }
                    }

                    string baseName = $"Temp_BlackOverrideFilter_Sheet_{sheet.Id.IntegerValue}";
                    string filterName = GetUniqueFilterName(doc, baseName);
                    ParameterFilterElement sheetFilter = ParameterFilterElement.Create(doc, filterName, categoryIdsToOverride);

                    tempFilterIds.Add(sheetFilter.Id);

                    OverrideGraphicSettings sheetOGS = CreateBlackOverrideSettings();

                    sheet.AddFilter(sheetFilter.Id);
                    sheet.SetFilterOverrides(sheetFilter.Id, sheetOGS);

                    // View-level filters
                    ICollection<ElementId> placedViewIds = sheet.GetAllPlacedViews();
                    foreach (ElementId viewId in placedViewIds)
                    {
                        Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                        if (view == null || view.IsTemplate) continue;

                        string viewBaseName = $"Temp_BlackOverrideFilter_View_{view.Id.IntegerValue}";
                        string viewFilterName = GetUniqueFilterName(doc, viewBaseName);
                        ParameterFilterElement viewFilter = ParameterFilterElement.Create(doc, viewFilterName, categoryIdsToOverride);

                        tempFilterIds.Add(viewFilter.Id);

                        OverrideGraphicSettings viewOGS = CreateBlackOverrideSettings();

                        view.AddFilter(viewFilter.Id);
                        view.SetFilterOverrides(viewFilter.Id, viewOGS);
                    }
                }

                tx.Commit();
            }

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

            //string outputFolder = @"C:\\Temp\\Revit Sheet PDFs";
            //Directory.CreateDirectory(outputFolder);

            string outputFolder = "";

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder to save printed PDFs";
                folderDialog.ShowNewFolderButton = true;

                DialogResult result = folderDialog.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                {
                    outputFolder = folderDialog.SelectedPath;
                }
                else
                {
                    TaskDialog.Show("Cancelled", "No folder selected. Printing aborted.");
                    return Result.Cancelled;
                }
            }


            List<string> printedFiles = new List<string>();

            foreach (ViewSheet sheet in sheetsCollector)
            {
                ViewSet vs = new ViewSet();
                vs.Insert(sheet);
                printManager.ViewSheetSetting.CurrentViewSheetSet.Views = vs;
                printManager.Apply();

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
                    int retry = 0;
                    while (!File.Exists(filePath) && retry < 10)
                    {
                        Thread.Sleep(500);
                        retry++;
                    }

                    if (File.Exists(filePath))
                        printedFiles.Add(filePath);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Print Error", $"Failed to print {sheet.Name}: {ex.Message}");
                    return Result.Failed;
                }
            }

            string mergedPdfPath = Path.Combine(outputFolder, "COMBINED_REVIT_SHEETS.pdf");
            MergePdfFiles(printedFiles, mergedPdfPath);

            using (Transaction cleanupTx = new Transaction(doc, "Clean Up Temporary Filters"))
            {
                cleanupTx.Start();
                foreach (ElementId filterId in tempFilterIds)
                {
                    try { doc.Delete(filterId); } catch { }
                }
                cleanupTx.Commit();
            }

            TaskDialog.Show("Success", $"Selected sheets printed and combined PDF saved to:\n{mergedPdfPath}");
            return Result.Succeeded;
        }

        private OverrideGraphicSettings CreateBlackOverrideSettings()
        {
            var ogs = new OverrideGraphicSettings();
            Color black = new Color(0, 0, 0);
            ogs.SetProjectionLineColor(black);
            ogs.SetCutLineColor(black);
            ogs.SetSurfaceBackgroundPatternColor(black);
            ogs.SetSurfaceForegroundPatternColor(black);
            ogs.SetCutBackgroundPatternColor(black);
            ogs.SetCutForegroundPatternColor(black);
            return ogs;
        }

        private string GetUniqueFilterName(Document doc, string baseName)
        {
            string name = baseName;
            int counter = 1;
            while (new FilteredElementCollector(doc).OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .Any(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                name = $"{baseName}_{counter}";
                counter++;
            }
            return name;
        }

        private void MergePdfFiles(List<string> filePaths, string outputPath)
        {
            PdfDocument outputDocument = new PdfDocument();

            foreach (string filePath in filePaths)
            {
                if (File.Exists(filePath))
                {
                    PdfDocument input = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                    for (int i = 0; i < input.PageCount; i++)
                    {
                        outputDocument.AddPage(input.Pages[i]);
                    }
                }
            }

            outputDocument.Save(outputPath);
        }
    }
}
