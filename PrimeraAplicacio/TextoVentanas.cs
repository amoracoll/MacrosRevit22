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


#endregion

namespace PrimeraAplicacio
{
    [Transaction(TransactionMode.Manual)]
    public class TextoVentanas : IExternalCommand
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
                //Obtenemos los puntos de las ventanas en questión:
                var puntos = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(FamilyInstance)).Where(x => x.Name == "0915 x 1830mm").
                    Select(lo => lo.Location).Cast<LocationPoint>().
                    Select(p => p.Point);

                //Encontram
                ElementId idTipo = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

                using(Transaction tran = new Transaction(doc,"Crear texto"))
                {
                    tran.Start();
                    foreach (var p in puntos)
                    {
                        XYZ mp = p + new XYZ(1, 2, 0);
                        TextNote texto = TextNote.Create(doc, doc.ActiveView.Id, mp, "Ventana a cambiar", idTipo);
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
