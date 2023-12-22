#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using OfficeOpenXml;


#endregion

namespace PrimeraAplicacio
{
    [Transaction(TransactionMode.Manual)]
    public class ExcelEscribir : IExternalCommand
    {
        //Para poder leer en Excel hay que tener instalado el Package PPlus y el namespace OfficeOpenXml
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            //Obtenemos el doc y el uidoc:
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                //Ruta del archivo Excel:
                string ruta = @"C:\Users\arnau.mora_modelical\Desktop\C#INCUBATOR\ProvaExcel.xlsx";

                //Para acceder a nuestro archivo Excel necesitamos crear un objeto de la clase FileInfo:
                var informacionArchivo = new FileInfo(ruta);

                //Creamos un objeto ExcelPackage:
                var paquete = new ExcelPackage(informacionArchivo);

                //Conseguimos su workseeet correspondiente (en este caso el 1):
                var worksheet = paquete.Workbook.Worksheets.Add("Planos actualizados");

                //Establecemos la primera fila con los encabezamientos:
                worksheet.SetValue(1, 1, "NÚMERO DE PLANO");
                worksheet.SetValue(1, 2, "NOMBRE DE PLANO");

                //Conseguimos todos los planos de nuestro proyecto:
                var planos = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>();

                int fila = 2;
                foreach(ViewSheet vS in planos)
                {
                    worksheet.SetValue(fila, 1, vS.SheetNumber);
                    worksheet.SetValue(fila, 2, vS.Name);
                    fila++;
                }
                paquete.Save();
                Process.Start(ruta);

                return Result.Succeeded;
            }
            catch(Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
                return Result.Failed;
            }
        }
    }
}
