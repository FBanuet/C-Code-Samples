using System;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using System.Runtime.InteropServices;
using Autodesk.Revit.UI.Selection;



namespace DATools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RevitToExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Selection sel = uidoc.Selection;
            ICollection<ElementId> eleids = sel.GetElementIds();

            using(Transaction trans = new Transaction(doc,"INFO"))
            {
                if(eleids.Count == 0)
                {
                    TaskDialog.Show("REVIT ERROR!", "NO HA SELECCIONADO NINGUN ELEMENTO AUN");
                }
                else
                {
                    Excel.Application xlApp = new Excel.Application();
                    if(xlApp == null)
                    {
                        TaskDialog.Show("EXCEL ERROR!", "EXCEL SE HA INSTALADO DE MANERA INCORRECTA!");
                        return Result.Failed;
                    }
                    Excel.Workbook xlworkbook;
                    Excel.Worksheet xlWorksheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlworkbook = xlApp.Workbooks.Add(misValue);
                    xlWorksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

                    //INTERACCIÓN CON REVIT Y SU DOCUMENTO

                    String info = "LOS IDS DE LOS ELEMENTOS EN EL DOCUMENTO SON : ";
                    int x = 1;
                    int y = 1;
                    xlWorksheet.Cells[x, y] = info;
                    x = x + 1;
                    foreach(ElementId item in eleids)
                    {
                        xlWorksheet.Cells[x, 1] = item.IntegerValue;
                        x = x + 1;
                    }

                    xlworkbook.SaveAs("D:\\Users\\S-MK\\Pictures\\test\\RevitTest.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive,
                        misValue, misValue, misValue, misValue, misValue);
                    xlworkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlworkbook);
                    Marshal.ReleaseComObject(xlApp);

                    TaskDialog.Show("REVIT", "EL EXCEL SE HA EXPORTADO EXITOSAMENTE!");

                }
            }

            return Result.Succeeded;
        }

        


    }
}
