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
using Autodesk.Revit.Creation;


#endregion

namespace PrimeraAplicacio
{
    [Transaction(TransactionMode.Manual)]
    public class ColocarElementosTxt : IExternalCommand
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
                string ruta = @"C:\Users\arnau.mora_modelical\Desktop\C#INCUBATOR\CreacionPlanos.xlsx";

                //Para acceder a nuestro archivo Excel necesitamos crear un objeto de la clase FileInfo:
                var informacionArchivo = new FileInfo(ruta);

                //Creamos un objeto ExcelPackage:
                var paquete = new ExcelPackage(informacionArchivo);

                //Conseguimos su workseeet correspondiente (en este caso el 1):
                var worksheet = paquete.Workbook.Worksheets[1];

                //Averiguamos el número de filas y columnas que tiene nuestro workseet:
                int filas = worksheet.Dimension.Rows;
                int columnas = worksheet.Dimension.Columns;

                //Creamos una lista que contenga las listas de cada fila:
                List<List<string>> listaFilas = new List<List<string>>();

                //Hacemos un foreachloop anidado:
                foreach(int i in Enumerable.Range(2,filas-1))
                {
                    //Creamos una lista para cada fila:
                    List<string> datosFilas = new List<string>();

                    foreach(int j in Enumerable.Range(1,columnas-1))
                    {
                        string valorNumero = worksheet.Cells[i,j].Value.ToString();
                        datosFilas.Add(valorNumero);
                    }
                    listaFilas.Add(datosFilas);
                }

                //Conseguimos el id del elemento a colocar:
                FamilySymbol tipoId = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).First(x => x.Name.Contains("A_PENDANT")).Id as FamilySymbol;


                //Empezamos una transacción:
                using (Transaction tran = new Transaction(doc,"Creación planos"))
                {
                    tran.Start();
                    foreach(var lF in listaFilas)
                    {
                        XYZ punto = new XYZ(lF[1], lF[2], lF[3]);
                        new FamilyInstanceCreationData(punto, tipoId, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    }
                    tran.Commit();
                }


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
