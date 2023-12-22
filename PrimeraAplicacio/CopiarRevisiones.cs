#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
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
    public class CopiarRevisiones : IExternalCommand
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

                //Filtramos las revisiones que queremos copiar:
                List<Revision> revisionRequerida = new FilteredElementCollector(doc).OfClass(typeof(Revision)).
                    Cast<Revision>().Where(x => x.RevisionNumber == "1").ToList();

                //Filtramos los planos que queremos:
                List<ViewSheet> planos = new FilteredElementCollector(doc).
                    OfClass(typeof(ViewSheet)).Cast<ViewSheet>().ToList();

                //Hacemos un foreach loop sobre cada plano para añadir la revisión:
                using (Transaction tran = new Transaction(doc,"Copiar revisiones"))
                {
                    tran.Start();
                    foreach(var p in planos)
                    {
                        var revisionesEnPlano = p.GetAdditionalRevisionIds();
                        foreach(var re in revisionRequerida)
                        {
                            if(!revisionesEnPlano.Contains<ElementId>(re.Id))
                            {
                                revisionesEnPlano.Add(re.Id);
                            }
                            p.SetAdditionalRevisionIds(revisionesEnPlano);
                        }
                        tran.Commit();
                    }
                }

                return Result.Succeeded;
            }
            catch(Exception e)
            {
                TaskDialog.Show("No se pudo copiar la revisión",e.ToString());
                return Result.Failed;
            }
        }
    }
}
