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
    public class CopiarVistaPlantilla : IExternalCommand
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
                //Seleccionamos las vistas que queremos duplicar:
                var seleccionIds = uidoc.Selection.GetElementIds();

                //So no hay nada seleccionado nos aparece un mensaje para que seleccionemos las vistas:
                if(seleccionIds.Count ==  0)
                {
                    TaskDialog.Show("Error en selección", "Por favor seleccione las vistas a duplicar con plantilla");
                }

                //Nos quedamos con la vista que será nuestra plantilla (la vista activa):
                ViewPlan vistaPlantilla = new FilteredElementCollector(doc).
                    OfClass(typeof(ViewPlan)).Cast<ViewPlan>().First(x => x.Name == "01 - Entry Level - Furniture Layout");

                using(Transaction tran = new Transaction(doc,"Copiamos vistas y aplicamos propiedades de vista"))
                {
                    tran.Start();
                    //Creamos una plantilla de vista a través de una vista que aplicaremos a posteriori:
                    ElementId plantilla = vistaPlantilla.CreateViewTemplate().Id;
                    foreach(var id in seleccionIds)
                    {
                        View vistaSeleccionada = doc.GetElement(id) as View;
                        ElementId vistaDuplicadaId = vistaSeleccionada.Duplicate(ViewDuplicateOption.Duplicate);
                        View vistaDuplicada = doc.GetElement(vistaDuplicadaId) as View;
                        //Aplicamos la plantilla creada anteriormente en las vistas duplicadas:
                        vistaDuplicada.ViewTemplateId = plantilla;
                        vistaDuplicada.Name = vistaSeleccionada.Name + "_con plantilla";
                    }
                }

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
