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

namespace RevitAPITrainingReadWrite
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string pipeInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .Cast<Pipe>()
                .ToList();

            string exelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(exelPath, FileMode.Create, FileAccess.Write)) 
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;

                foreach (var pipe in pipes)
                {
                    sheet.SetCellValue(rowIndex, columIndex: 0, pipe.Name);
                    sheet.SetCellValue(rowIndex, columIndex: 1, pipe.LookupParameter("Внутренний диаметр"));
                    sheet.SetCellValue(rowIndex, columIndex: 2, pipe.LookupParameter("Внешний диаметр"));
                    sheet.SetCellValue(rowIndex, columIndex: 3, pipe.LookupParameter("Длина"));
                    rowIndex++;
                }

                workbook.Write(stream);
                workbook.Close();
            }
            System.Diagnostics.Process.Start(exelPath);

            return Result.Succeeded;
        }
    }
}
