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
    public class PrimerFiltroVista : IExternalCommand
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
                //Seleccionamos el Id de la categoría que me interesa:
                ElementId idPuertas = new elementid(BuiltInCategory.OST_Doors);
                List<ElementId> categorias = new List<ElementId>();
                categorias.Add(idPuertas);

                //Creamos el filter rule utilizando el nombre del tipo de puerta:
                FilterRule fR = ParameterFilterRuleFactory.CreateEqualsRule(new ElementId(BuiltInParameter.SYMBOL_NAME_PARAM), "90 x 210 cm", false);

                //Creamos el filtro (ElementPa5rameterFilter):
                ElementParameterFilter filtro = new ElementParameterFilter(fR);

                //Iniciamos transacción:
                using(Transaction tran = new Transaction(doc,"Aplicamos el filtro"))
                {
                    tran.Start();
                    //Creamos el parameter filter element:
                    ParameterFilterElement paFiEle = ParameterFilterElement.Create(doc,"puertas",categorias, filtro);

                    //Añadimos el filtro a la vista:
                    doc.ActiveView.AddFilter(paFiEle.Id);

                    //Establecemos la visibilidad:
                    doc.ActiveView.SetFilterVisibility(paFiEle.Id, false);

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
