#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows.Media.Imaging;

#endregion

namespace PrimeraAplicacio
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            //Creación de la botonera:
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().Location;

            RibbonPanel panel = a.CreateRibbonPanel("Mi primera aplicación");

            PushButtonData BotonSaludar = new PushButtonData("Botón1", "MAJA",assemblyName, "PrimeraAplicacio.Command");

            PushButton boton1 = panel.AddItem(BotonSaludar) as PushButton;

            boton1.ToolTip = "Modelicos enviats a Shushah";
            boton1.LongDescription = "Arnau M | Arnau P | Jordi | Maria";
            boton1.LargeImage = new BitmapImage(new Uri(@"C:\Users\arnau.mora_modelical\Downloads\icons8-saludo-ios-16\icons8-mierda-32.png"));

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
