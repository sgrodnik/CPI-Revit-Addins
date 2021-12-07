using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using FavoriteViews;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CPI {
    [Transaction(TransactionMode.Manual)]
    public class FavoriteViews : IExternalCommand {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            MainWindow window = new MainWindow(commandData);
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(window);
            helper.Owner = uiApp.MainWindowHandle;
            window.Title = "Избранные виды";
            window.Show();

            return Result.Succeeded;

        }
    }
}
