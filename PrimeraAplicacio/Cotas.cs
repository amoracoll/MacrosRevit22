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
    public class Cotas : IExternalCommand
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

                //Obtenemos todas las cotas distintas de Diagonal:
                var cotas = new FilteredElementCollector(doc).OfClass(typeof(Dimension)).Cast<Dimension>().
                    Where(x => x.DimensionType.Name != "Diagonal" && x.DimensionType.StyleType == DimensionStyleType.Linear);

                //Creamos un texto de aviso por encima a las cotas que no se corresponden con Diagonal:
                using(Transaction tran = new Transaction(doc,"Colocar comentario por encima"))
                {
                    tran.Start();
                    foreach(var cota in cotas)
                    {
                        var numeroSegmentos = cota.NumberOfSegments;
                        if(numeroSegmentos == 0) 
                        {
                            cota.Above = "Hay que cambiar estilo de cota";
                        }
                        else
                        {
                            foreach(var numero in Enumerable.Range(0,numeroSegmentos))
                            {
                                var segmentos = cota.Segments;
                                segmentos.get_Item(numero).Above = "Hay que cambiar estilo de cota";
                            }
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
