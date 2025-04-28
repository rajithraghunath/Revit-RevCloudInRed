using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ForceHatchBlackOnSelectedSheets : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Step 1: Collect Selected Sheets
        IList<ElementId> selectedIds = uidoc.Selection.GetElementIds().ToList();
        if (selectedIds.Count == 0)
        {
            TaskDialog.Show("Warning", "Please select sheet(s) before running this command.");
            return Result.Failed;
        }

        List<ViewSheet> selectedSheets = new List<ViewSheet>();
        foreach (ElementId id in selectedIds)
        {
            Element elem = doc.GetElement(id);
            if (elem is ViewSheet sheet)
            {
                selectedSheets.Add(sheet);
            }
        }

        if (selectedSheets.Count == 0)
        {
            TaskDialog.Show("Warning", "No sheets selected.");
            return Result.Failed;
        }

        // Step 2: Set up override (black hatch patterns)
        OverrideGraphicSettings blackOverride = new OverrideGraphicSettings();
        Color blackColor = new Color(0, 0, 0);
        blackOverride.SetCutForegroundPatternColor(blackColor);
        blackOverride.SetCutBackgroundPatternColor(blackColor);
        blackOverride.SetProjectionLineColor(blackColor);
        blackOverride.SetProjectionLinePatternId(ElementId.InvalidElementId); // No pattern

        //blackOverride.SetProjectionForegroundPatternColor(blackColor);
        //blackOverride.SetProjectionBackgroundPatternColor(blackColor);

        // Step 3: Collect Categories to override
        List<BuiltInCategory> categoriesToOverride = new List<BuiltInCategory>()
        {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_FilledRegion
            // Add more categories if needed
        };

        // Step 4: Start Transaction - Apply Overrides
        using (Transaction tx = new Transaction(doc, "Apply Black Hatch Overrides"))
        {
            tx.Start();
            foreach (ViewSheet sheet in selectedSheets)
            {
                foreach (BuiltInCategory bic in categoriesToOverride)
                {
                    Category cat = Category.GetCategory(doc, bic);
                    if (cat != null)
                    {
                        sheet.SetCategoryOverrides(cat.Id, blackOverride);
                    }
                }
            }
            tx.Commit();
        }

        // Step 5: Open Print Dialog
        // Fix: Replace the incorrect `ShowPrintDialog` method with the correct approach to invoke the print dialog.
        PrintManager printManager = doc.PrintManager;
        printManager.SelectNewPrintDriver("Microsoft Print to PDF"); // Replace with your desired printer name
        printManager.PrintRange = PrintRange.Select;
        printManager.Apply();

        // Use the Revit API's PrintManager to handle printing instead of a non-existent `ShowPrintDialog` method.
        TaskDialog.Show("Print", "Please open the print dialog manually to proceed with printing.");

        // Step 6: Start Transaction - Reset Overrides After Printing
        using (Transaction tx = new Transaction(doc, "Reset Hatch Overrides"))
        {
            tx.Start();
            OverrideGraphicSettings resetOverride = new OverrideGraphicSettings(); // blank settings
            foreach (ViewSheet sheet in selectedSheets)
            {
                foreach (BuiltInCategory bic in categoriesToOverride)
                {
                    Category cat = Category.GetCategory(doc, bic);
                    if (cat != null)
                    {
                        sheet.SetCategoryOverrides(cat.Id, resetOverride);
                    }
                }
            }
            tx.Commit();
        }

        return Result.Succeeded;
    }
}
