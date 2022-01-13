using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RvtAPI_HW4_2_PipesOutput
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Pipe> pipesInstances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var pipeInstance in pipesInstances)
                {
                    string pipeName = pipeInstance.Name;
                    double outsideDiamParam = pipeInstance.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                    double outsideDiamMm = UnitUtils.ConvertFromInternalUnits(outsideDiamParam, UnitTypeId.Millimeters);
                    double insideDiamParam = pipeInstance.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                    double insideDiamMm = UnitUtils.ConvertFromInternalUnits(insideDiamParam, UnitTypeId.Millimeters);
                    double lengthParam = pipeInstance.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    double lengthMm = UnitUtils.ConvertFromInternalUnits(lengthParam, UnitTypeId.Millimeters);
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipeName);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outsideDiamMm);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, insideDiamMm);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, lengthMm);
                    rowIndex++;
                }
                workBook.Write(stream);
                workBook.Close();
            }
            TaskDialog.Show("Запись параметров в текстовый файл", "Запись завершена");
            return Result.Succeeded;
        }
    }
}
