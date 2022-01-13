using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RvtAPI_HW4_1_WallsOutput
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .ToList();

            string wallInfo = string.Empty;
            foreach (Wall wall in walls)
            {
                double wallVolumeFeet = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble();
                string wallVolumeMeters = UnitUtils.ConvertFromInternalUnits(wallVolumeFeet, UnitTypeId.CubicMeters).ToString();
                wallInfo += $"{wall.Name}\t{"|"}\t{ wallVolumeMeters}{Environment.NewLine}";
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string csvPath = Path.Combine(desktopPath, "wallInfo.csv");

            File.WriteAllText(csvPath, wallInfo);

            TaskDialog.Show("Запись параметров в текстовый файл", "Запись завершена");

            return Result.Succeeded;
        }
    }
}
