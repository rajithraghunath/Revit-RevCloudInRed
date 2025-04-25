using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevCloudinRed
{
    class CutMaterial
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Pick a wall
            Reference pickedRef = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            Wall wall = doc.GetElement(pickedRef) as Wall;

            //Reference allRef = uiDoc.Document.ActiveView.GetReferenceByElementId(wall.Id);

            if (wall == null)
            {
                message = "Selected element is not a wall.";
                return Result.Failed;
            }

            using (Transaction trans = new Transaction(doc, "Change Cut Hatch Pattern"))
            {
                trans.Start();

                WallType wallType = wall.WallType;
                CompoundStructure compStruct = wallType.GetCompoundStructure();

                // Get the material from the first structural layer
                IList<CompoundStructureLayer> layers = compStruct.GetLayers();
                Material material = null;

                foreach (var layer in layers)
                {
                    if (layer.Function == MaterialFunctionAssignment.Structure)
                    {
                        material = doc.GetElement(layer.MaterialId) as Material;
                        break;
                    }
                }

                if (material == null)
                {
                    message = "No structural material found.";
                    return Result.Failed;
                }

                // Load or find a pattern (e.g., "Diagonal Crosshatch")
                FillPatternElement patternElement = new FilteredElementCollector(doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .FirstOrDefault(p => p.GetFillPattern().Name == "Diagonal Crosshatch");

                if (patternElement == null)
                {
                    message = "Pattern not found.";
                    return Result.Failed;
                }

                // Apply the pattern and color to the material’s surface pattern

                //material.SurfacePatternId = patternElement.Id;
                //material.SurfaceColor = new Color(0, 0, 0); // Black

                material.CutBackgroundPatternColor = new Color(255, 0, 0); // Red
                material.CutForegroundPatternColor = new Color(255, 0, 0); // Red
                material.SurfaceBackgroundPatternColor = new Color(255, 0, 0); // Red
                material.SurfaceForegroundPatternColor = new Color(255, 0, 0); // Red


                //material.ForegroundPatternId = patternElement.Id;
                //material.ForegroundColor = new Color(0, 0, 0); // Black

                trans.Commit();
            }

            return Result.Succeeded;
        }

    }
}
