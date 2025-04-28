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

                    BuiltInCategory[] mustInclude = new[]
                    {
                        BuiltInCategory.OST_CutOutlines,
                        BuiltInCategory.OST_Doors,

                        BuiltInCategory.OST_Materials,

                        BuiltInCategory.OST_Rooms,

                        BuiltInCategory.OST_FillPatterns,
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_WallsCutPattern,
                        BuiltInCategory.OST_WallsDefault,
                        BuiltInCategory.OST_WallsFinish1, BuiltInCategory.OST_WallsFinish2,
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


                    //**********
                    //Base Class: SheetFilter
                    #region SheetFilter
                    ParameterFilterElement sheetFilter = ParameterFilterElement.Create(doc,
                        $"Temp_BlackOverrideFilter_Sheet_{sheet.Id.IntegerValue}",
                        categoryIdsToOverride);
                    tempFilterIds.Add(sheetFilter.Id);

                    OverrideGraphicSettings sheetOGS = new OverrideGraphicSettings();
                    sheetOGS.SetProjectionLineColor(new Color(0, 0, 0));
                    sheetOGS.SetCutLineColor(new Color(0, 0, 0));
                    sheetOGS.SetSurfaceForegroundPatternColor(new Color(0, 0, 0));
                    sheetOGS.SetCutBackgroundPatternColor(new Color(0, 0, 0));
                    
                    sheet.AddFilter(sheetFilter.Id);
                    sheet.SetFilterOverrides(sheetFilter.Id, sheetOGS);

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

                        #endregion

                        #region WallHatchPattern
                        // 🔍 Collect wall cut hatch patterns
                        List<Wall> wallCollector = new FilteredElementCollector(doc, view.Id)

                            .OfClass(typeof(Wall))
                            .Cast<Wall>()
                            .Where(w => w != null)
                            .ToList();

                        foreach (Wall wall in wallCollector)
                        {
                            WallType wallType = doc.GetElement(wall.GetTypeId()) as WallType;
                            if (wallType == null) continue;

                            Material material = null;
                            FillPatternElement cutPattern = null;

                            CompoundStructure cs = wallType.GetCompoundStructure();
                            if (cs != null)
                            {
                                foreach (CompoundStructureLayer layer in cs.GetLayers())
                                {
                                    if (layer.Function == MaterialFunctionAssignment.Structure)
                                    {
                                        material = doc.GetElement(layer.MaterialId) as Material;
                                        break;
                                    }
                                }
                            }

                            if (material != null)
                            {
                                ElementId cutPatternId = material.CutForegroundPatternId;
                                if (cutPatternId != ElementId.InvalidElementId)
                                {
                                    cutPattern = doc.GetElement(cutPatternId) as FillPatternElement;
                                    if (cutPattern != null)
                                    {
                                        string patternName = cutPattern.Name;
                                        TaskDialog.Show("Wall Cut Pattern", $"View: {view.Name}\nWall: {wall.Id}\nPattern: {patternName}");
                                    }
                                }
                            }

                            #endregion

                        }
                    }
                }

                tx.Commit();
            }

            #region SheetManager

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
            Directory.CreateDirectory(outputFolder);

            List<string> printedFiles = new List<string>();

            foreach (ViewSheet sheet in sheetsCollector)
            {
                ViewSet vs = new ViewSet();
                vs.Insert(sheet);
                printManager.ViewSheetSetting.CurrentViewSheetSet.Views = vs;
                printManager.Apply();

                string fileName = $"{sheet.SheetNumber}_{sheet.Name}.pdf";
                foreach (char c in Path.GetInvalidFileNameChars())
                    fileName = fileName.Replace(c, '_');

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
            #endregion

            string mergedPdfPath = Path.Combine(outputFolder, "COMBINED_REVIT_SHEETS.pdf");
            MergePdfFiles(printedFiles, mergedPdfPath);


            //************
            //************
            //************

            using (Transaction cleanupTx = new Transaction(doc, "Clean Up Temporary Filters"))
            {
                cleanupTx.Start();
                foreach (ElementId filterId in tempFilterIds)
                {
                    try { doc.Delete(filterId); } catch { }
                }
                cleanupTx.Commit();
            }

            TaskDialog.Show("Success", $"All sheets printed and combined PDF saved to:\n{mergedPdfPath}");
            return Result.Succeeded;

        }

        #region PDF Merge
        // Merge the printed PDF files
        private void MergePdfFiles(List<string> filePaths, string outputPath)
        {
            PdfDocument outputDocument = new PdfDocument();
            foreach (string filePath in filePaths)
            {
                if (File.Exists(filePath))
                {
                    PdfDocument input = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
                    for (int i = 0; i < input.PageCount; i++)
                        outputDocument.AddPage(input.Pages[i]);
                }
            }
            outputDocument.Save(outputPath);
        }

        #endregion

    }
}
