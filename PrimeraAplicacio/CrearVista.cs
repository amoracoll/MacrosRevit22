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

#endregion

namespace PrimeraAplicacio
{
    [Transaction(TransactionMode.Manual)]
    public class CrearVista : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            //Obtenemos el doc:
            UIDocument uidoc = commandData.Application.ActiveUIDocument;    //Canviar this por commandData.Application
            Document doc = uidoc.Document;

            try
            {
                //Determinamos la altura del nivel:
                double altura = UnitUtils.ConvertToInternalUnits(8, UnitTypeId.Meters);

                //Conseguimos el id de FloorPlan:
                ElementId floorPlanId = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                .First(x => x.ViewFamily == ViewFamily.FloorPlan).Id;

                //Iniciamos transacción:
                using (Transaction tran = new Transaction(doc, "Creamos una vista"))
                {
                    tran.Start();
                    //Creamos nivel:
                    Level nuevoNivel = Level.Create(doc, altura);
                    nuevoNivel.Name = "Nivel Prueba";

                    //Creamos vista:
                    ViewPlan nuevaVista = ViewPlan.Create(doc, floorPlanId, nuevoNivel.Id);
                    nuevaVista.Name = "Nivel Prueba_vista_01";
                    nuevaVista.Scale = 200;

                    tran.Commit();
                }
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
            }
            return Result.Succeeded;
        }
    }
}
