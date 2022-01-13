using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_CW4_5
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string roomInfo = string.Empty;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();
            
            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var room in rooms)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, room.Number);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, room.Area);
                    rowIndex++;
                }

                workBook.Write(stream);
                workBook.Close();
            }

            System.Diagnostics.Process.Start(excelPath);

            return Result.Succeeded;
        }
    }
}
