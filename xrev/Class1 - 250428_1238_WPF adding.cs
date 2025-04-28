using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Forms;

namespace RevCloudInRed
{
    [Transaction(TransactionMode.Manual)]
    public class PrintWithColoredRevisionClouds : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Show sheet picker window
            var sheetPickerWindow = new SheetPickerWindow(doc);

            // Safely invoke the dialog in the Revit thread context
            bool? dialogResult = null;
            object value = commandData.Application.Application.Invoke(() =>
            {
                dialogResult = sheetPickerWindow.ShowDialog();
            });

            if (dialogResult == null || !dialogResult.Value || sheetPickerWindow.SelectedSheets.Count == 0)
            {
                TaskDialog.Show("Warning", "No sheets selected.");
                return Result.Cancelled;
            }

            List<ViewSheet> sheetsToPrint = sheetPickerWindow.SelectedSheets;

            List<ElementId> tempFilterIds = new List<ElementId>();

            using (Transaction tx = new Transaction(doc, "Create Black Override Filters"))
            {
                tx.Start();

                foreach (ViewSheet sheet in sheetsToPrint)
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

                    // Include key categories
                    BuiltInCategory[] mustInclude = new[] { ... }; // Same as before

                    foreach (BuiltInCategory bic in mustInclude)
                    {
                        Category extraCat = Category.GetCategory(doc, bic);
                        if (extraCat != null && !categoryIdsToOverride.Contains(extraCat.Id))
                        {
                            categoryIdsToOverride.Add(extraCat.Id);
                        }
                    }

                    // Create sheet-level filter
                    string baseName = $"Temp_BlackOverrideFilter_Sheet_{sheet.Id.IntegerValue}";
                    string filterName = GetUniqueFilterName(doc, baseName);

                    ParameterFilterElement sheetFilter = ParameterFilterElement.Create(doc, filterName, categoryIdsToOverride);
                    tempFilterIds.Add(sheetFilter.Id);

                    OverrideGraphicSettings sheetOGS = CreateBlackOverride();
                    sheet.AddFilter(sheetFilter.Id);
                    sheet.SetFilterOverrides(sheetFilter.Id, sheetOGS);

                    // View-level filters (same as before)
                    ICollection<ElementId> placedViewIds = sheet.GetAllPlacedViews();
                    foreach (ElementId viewId in placedViewIds)
                    {
                        View view = doc.GetElement(viewId) as View;
                        if (view == null || view.IsTemplate) continue;

                        string viewBaseName = $"Temp_BlackOverrideFilter_View_{view.Id.IntegerValue}";
                        string viewFilterName = GetUniqueFilterName(doc, viewBaseName);

                        ParameterFilterElement viewFilter = ParameterFilterElement.Create(doc, viewFilterName, categoryIdsToOverride);
                        tempFilterIds.Add(viewFilter.Id);

                        OverrideGraphicSettings viewOGS = CreateBlackOverride();
                        view.AddFilter(viewFilter.Id);
                        view.SetFilterOverrides(viewFilter.Id, viewOGS);
                    }
                }

                tx.Commit();
            }

            // Print logic remains the same...
            // (The rest of the print process remains unchanged)
            return Result.Succeeded;
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

        private OverrideGraphicSettings CreateBlackOverride()
        {
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            Color black = new Color(0, 0, 0);
            ogs.SetProjectionLineColor(black);
            ogs.SetCutLineColor(black);
            ogs.SetSurfaceBackgroundPatternColor(black);
            ogs.SetSurfaceForegroundPatternColor(black);
            ogs.SetCutBackgroundPatternColor(black);
            ogs.SetCutForegroundPatternColor(black);
            return ogs;
        }
    }

    public class SheetPickerWindow : Window
    {
        public List<ViewSheet> SelectedSheets { get; private set; }
        private ListBox sheetListBox;

        public SheetPickerWindow(Document doc)
        {
            Title = "Select Sheets to Print";
            Width = 300;
            Height = 400;
            SelectedSheets = new List<ViewSheet>();

            // Create ListBox for sheet selection
            sheetListBox = new ListBox
            {
                SelectionMode = SelectionMode.Multiple,
                Margin = new Thickness(10)
            };

            // Populate ListBox with sheets
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .WhereElementIsNotElementType();

            foreach (ViewSheet sheet in collector)
            {
                sheetListBox.Items.Add(new ListBoxItem { Content = $"{sheet.SheetNumber} - {sheet.Name}", Tag = sheet });
            }

            // Add Select and Cancel buttons
            Button selectButton = new Button { Content = "Select", Margin = new Thickness(10) };
            selectButton.Click += (s, e) =>
            {
                foreach (ListBoxItem item in sheetListBox.SelectedItems)
                {
                    ViewSheet sheet = item.Tag as ViewSheet;
                    if (sheet != null)
                    {
                        SelectedSheets.Add(sheet);
                    }
                }

                DialogResult = true;
                Close();
            };

            Button cancelButton = new Button { Content = "Cancel", Margin = new Thickness(10) };
            cancelButton.Click += (s, e) => Close();

            StackPanel panel = new StackPanel();
            panel.Children.Add(sheetListBox);
            panel.Children.Add(selectButton);
            panel.Children.Add(cancelButton);

            Content = panel;
        }
    }
}
