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
    public class ExcelLeer : IExternalCommand
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
                var worksheet = paquete.Workbook.Worksheets[1];

                //Averiguamos el número de filas y columnas que tiene nuestro workseet:
                int filas = worksheet.Dimension.Rows;
                int columnas = worksheet.Dimension.Columns;

                //Hacemos un foreachloop anidado:
                string datos = "Los valores de cada fila son: " + "\n\n";
                foreach(int i in Enumerable.Range(2,filas-1))
                {
                    foreach(int j in Enumerable.Range(1,columnas-1))
                    {
                        string valor = worksheet.Cells[i, j].Value.ToString();
                        datos += valor + " ";
                    }
                    datos += "\n";
                }
                TaskDialog.Show("Contenido Excel", datos);


                return Result.Succeeded;
            }
            catch(Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}
