using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;

namespace RevCloudInRed
{
    public class SheetPickerWindow : Window
    {
        public List<ViewSheet> SelectedSheets { get; private set; }
        private ListBox sheetListBox;

        // Ensure this is the only constructor for this class
        public SheetPickerWindow(Document doc)
        {
            // Initialize window properties
            Title = "Select Sheets to Print";
            Width = 300;
            Height = 400;
            SelectedSheets = new List<ViewSheet>();

            // Create ListBox for sheet selection
            ListBox sheetListBox = new ListBox
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

            // Set up StackPanel with ListBox and buttons
            StackPanel panel = new StackPanel();
            panel.Children.Add(sheetListBox);
            panel.Children.Add(selectButton);
            panel.Children.Add(cancelButton);

            Content = panel;
        }
    }
}
