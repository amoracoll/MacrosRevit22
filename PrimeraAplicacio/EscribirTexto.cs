#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

#endregion

namespace PrimeraAplicacio
{
    [Transaction(TransactionMode.Manual)]
    public class EscribirTexto : IExternalCommand
    {
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
                //Creamos un ejemplar de StreamWriter utilizando using:
                string ruta = @"C:\Users\arnau.mora_modelical\Desktop\MisComandos\otro documento.txt";
                using (StreamWriter escritor = new StreamWriter(ruta,false))
                {
                    escritor.WriteLine("\nnueva línea de código");
                    escritor.WriteLine("otra línea");
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
