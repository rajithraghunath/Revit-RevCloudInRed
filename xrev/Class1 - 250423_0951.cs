using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
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

            List<ViewSheet> sheetsCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(sheet => !sheet.IsPlaceholder)
                .ToList();

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

                    // Explicitly include key categories
                    BuiltInCategory[] mustInclude = new[]
                    {
                        BuiltInCategory.OST_Rooms,
                        BuiltInCategory.OST_FillPatterns,
                        //BuiltInCategory.OST_ColorFillLegends,
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_WallsCutPattern,
                        BuiltInCategory.OST_Doors,

                        BuiltInCategory.OST_WallsSurfacePattern,
                        
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

                    // Create sheet-level filter
                    ParameterFilterElement sheetFilter = ParameterFilterElement.Create(doc,
                        $"Temp_BlackOverrideFilter_Sheet_{sheet.Id.IntegerValue}",
                        categoryIdsToOverride);
                    tempFilterIds.Add(sheetFilter.Id);

                    OverrideGraphicSettings sheetOGS = new OverrideGraphicSettings();
                    sheetOGS.SetProjectionLineColor(new Color(0, 0, 0));
                    sheetOGS.SetCutLineColor(new Color(0, 0, 0));
                    sheetOGS.SetSurfaceForegroundPatternColor(new Color(0, 0, 0));

                    sheetOGS.SetCutBackgroundPatternColor(new Color(0, 0, 0));
                    sheetOGS.SetCutBackgroundPatternColor(new Color(0, 0, 0));

                    sheet.AddFilter(sheetFilter.Id);
                    sheet.SetFilterOverrides(sheetFilter.Id, sheetOGS);

                    // View-level filters
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
                        viewOGS.SetCutLineColor(new Color(0, 0, 0));
                        viewOGS.SetSurfaceForegroundPatternColor(new Color(0, 0, 0));

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

            string outputFolder = @"C:\Temp\Revit Sheet PDFs";
            Directory.CreateDirectory(outputFolder); // Ensure the output directory exists

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
                    Directory.CreateDirectory(fileName); // Ensure the File exists

                }


                string filePath = Path.Combine(outputFolder, fileName);
                printManager.PrintToFileName = filePath;

                try
                {
                    printManager.SubmitPrint();

                    // Wait and retry to ensure the file is created
                    int retry = 0;
                    while (!File.Exists(filePath) && retry < 10)
                    {
                        Thread.Sleep(500); // Wait 0.5 seconds
                        retry++;
                    }

                    if (File.Exists(filePath))
                        printedFiles.Add(filePath);
                    //else
                    //    TaskDialog.Show("Warning", $"File not created: {filePath}");

                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Print Error", $"Failed to print {sheet.Name}: {ex.Message}");
                    return Result.Failed;
                }
            }

            // Merge PDFs into single file
            string mergedPdfPath = Path.Combine(outputFolder, "COMBINED_REVIT_SHEETS.pdf");
            MergePdfFiles(printedFiles, mergedPdfPath);

            using (Transaction cleanupTx = new Transaction(doc, "Clean Up Temporary Filters"))
            {
                cleanupTx.Start();
                foreach (ElementId filterId in tempFilterIds)
                {
                    try
                    {
                        doc.Delete(filterId);
                    }
                    catch { }
                }
                cleanupTx.Commit();
            }

            TaskDialog.Show("Success", $"All sheets printed and combined PDF saved to:\n{mergedPdfPath}");
            return Result.Succeeded;
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
