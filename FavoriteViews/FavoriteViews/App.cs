using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System.Reflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CPI {
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication {
        public Result OnShutdown(UIControlledApplication application) {
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application) {
            AddRibbonPanel(application);
            return Result.Succeeded;
        }
        static void AddRibbonPanel(UIControlledApplication application) {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Test");
            string path = Assembly.GetExecutingAssembly().Location;
            PushButtonData pbData = new PushButtonData("cmdTest", "Test", path, "CPI.FavoriteViews");
            PushButton pb = ribbonPanel.AddItem(pbData) as PushButton;
        }
    }
}