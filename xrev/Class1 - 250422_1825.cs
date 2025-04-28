using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

            // Get selected sheets only
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            List<ViewSheet> sheetsCollector = new FilteredElementCollector(doc)
             .OfClass(typeof(ViewSheet))
             .Cast<ViewSheet>()
             .Where(sheet => !sheet.IsPlaceholder)
             .ToList();

            if (sheetsCollector.Count == 0)
            {
                TaskDialog.Show("No Sheets Selected", "Please select sheet views before running the command.");
                return Result.Cancelled;
            }

            List<ElementId> tempFilterIds = new List<ElementId>();

            // Create black override filters for sheets and views
            using (Transaction tx = new Transaction(doc, "Create Black Override Filters"))
            {
                tx.Start();

                foreach (ViewSheet sheet in sheetsCollector)
                {
                    // Collect model & annotation categories except Revision Clouds, with forced override on specific categories
                    List<ElementId> categoryIdsToOverride = new List<ElementId>();

                    List<BuiltInCategory> forceBlackCategories = new List<BuiltInCategory>
                    {
                        BuiltInCategory.OST_Rooms,
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_Windows,
                        BuiltInCategory.OST_Doors,
                        BuiltInCategory.OST_WallsCutPattern
                    };

                    foreach (Category cat in doc.Settings.Categories)
                    {
                        if ((cat.CategoryType == CategoryType.Annotation || cat.CategoryType == CategoryType.Model) &&
                            cat.Id.IntegerValue != (int)BuiltInCategory.OST_RevisionClouds)
                        {
                            categoryIdsToOverride.Add(cat.Id);
                        }

                        if (forceBlackCategories.Contains((BuiltInCategory)cat.Id.IntegerValue) &&
                            !categoryIdsToOverride.Contains(cat.Id))
                        {
                            categoryIdsToOverride.Add(cat.Id);
                        }
                    }

                    // Sheet-level filter
                    ParameterFilterElement sheetFilter = ParameterFilterElement.Create(doc,
                        $"Temp_BlackOverrideFilter_Sheet_{sheet.Id.IntegerValue}",
                        categoryIdsToOverride);
                    tempFilterIds.Add(sheetFilter.Id);

                    OverrideGraphicSettings sheetOGS = new OverrideGraphicSettings();
                    sheetOGS.SetProjectionLineColor(new Color(0, 0, 0));

                    sheet.AddFilter(sheetFilter.Id);
                    sheet.SetFilterOverrides(sheetFilter.Id, sheetOGS);

                    // View-level filters for placed views on the sheet
                    ICollection<ElementId> placedViewIds = sheet.GetAllPlacedViews();
                    foreach (ElementId viewId in placedViewIds)
                    {
                        View view = doc.GetElement(viewId) as View;
                        if (view == null || view.IsTemplate) continue;

                        ParameterFilterElement viewFilter = ParameterFilterElement.Create(doc,
                            $"Temp_BlackOverrideFilter_View_{view.Id.IntegerValue}",
                            categoryIdsToOverride);
                        tempFilterIds.Add(viewFilter.Id);

                        OverrideGraphicSettings viewOGS = new OverrideGraphicSettings();
                        viewOGS.SetProjectionLineColor(new Color(0, 0, 0));

                        view.AddFilter(viewFilter.Id);
                        view.SetFilterOverrides(viewFilter.Id, viewOGS);
                    }
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

                parameters.ColorDepth = ColorDepthType.Color; // Ensure revision clouds appear in color
                parameters.PaperPlacement = PaperPlacementType.Center;
                parameters.ZoomType = ZoomType.FitToPage;
                parameters.HideCropBoundaries = true;

                printTx.Commit();
            }

            // Output folder
            string outputFolder = @"C:\Temp\Revit Sheet PDFs";
            Directory.CreateDirectory(outputFolder);

            List<string> pdfFiles = new List<string>();

            // Print each selected sheet
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
                    pdfFiles.Add(filePath);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Print Error", $"Failed to print {sheet.Name}: {ex.Message}");
                    return Result.Failed;
                }
            }

            // Merge all PDF files into one
            string mergedPath = Path.Combine(outputFolder, "CombinedSheets.pdf");
            using (PdfDocument outputDoc = new PdfDocument())
            {
                foreach (string pdf in pdfFiles)
                {
                    using (PdfDocument inputDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < inputDoc.PageCount; i++)
                        {
                            outputDoc.AddPage(inputDoc.Pages[i]);
                        }
                    }
                }
                outputDoc.Save(mergedPath);
            }

            // Clean up temporary filters
            using (Transaction cleanupTx = new Transaction(doc, "Clean Up Temporary Filters"))
            {
                cleanupTx.Start();
                foreach (ElementId filterId in tempFilterIds)
                {
                    try
                    {
                        doc.Delete(filterId);
                    }
                    catch { /* Ignore if already deleted */ }
                }
                cleanupTx.Commit();
            }

            TaskDialog.Show("Success", $"All selected sheets printed and merged into:\n{mergedPath}");
            return Result.Succeeded;
        }
    }
}
