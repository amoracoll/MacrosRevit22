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
    public class CrearEtiquetas : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            try
            {
                //Obtenemos el doc y el uidoc:
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                //Filtramos las vistas que contienen la palabra tags:
                var vistas = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
                     Where(x => x.Name.Contains("tags"));

                //Filtramos las puertas y ventanas del proyecto:
                var puertasVentanas = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).
                    Where(x => x.Category.Name == "Doors" || x.Category.Name == "Windows");

                using(Transaction tran = new Transaction(doc,"Crear etiqueta"))
                {
                    tran.Start();
                    foreach (var vista in vistas)
                    {
                        foreach(var elem in puertasVentanas)
                        {
                            XYZ punto = (elem.Location as LocationPoint).Point;
                            Reference referencia = new Reference(elem);
                            IndependentTag.Create(doc, vista.Id, referencia, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, punto);
                        }
                    }
                    tran.Commit();
                }

                return Result.Succeeded;

            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
                return Result.Failed;
            }
        }
    }
}
