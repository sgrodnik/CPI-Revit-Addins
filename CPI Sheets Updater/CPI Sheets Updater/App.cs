using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace CPI_Sheets_Updater {
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication {
        public static Dictionary<Document, bool> UpdaterIsAvailable = new Dictionary<Document, bool>();
        private static string[] paramNames = {
                "CPI_Количество листов",
                "CPI_ВС Наименование спецификации",
                "CPI_Группирование видов",
                "ADSK_Назначение вида",
                "CPI_Номер листа",
                "CPI_ВЧ Номер листа",
                "CPI_ВЧ Примечание",
                "CPI_ВС Номер листа",
                "CPI_ВС Примечание",
            };
        public static PushButton button;
        public static bool UpdaterIsEnabledFlag = true;

        static void AddRibbonPanel(UIControlledApplication application) {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("53 ЦПИ");
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData pbData = new PushButtonData(
                "cmdSwitchUpdater",
                "Остановить" + Environment.NewLine + "Sheets Updater",
                thisAssemblyPath,
                "CPI_Sheets_Updater.SwitchUpdater");
            button = ribbonPanel.AddItem(pbData) as PushButton;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://bit.do/Sheets-Updater"));
            buttonOn();
        }

        public static void buttonOn() {
            button.LargeImage = new BitmapImage(new Uri("pack://application:,,,/CPI Sheets Updater;component/Resources/toggle-on.png"));
            button.Image = new BitmapImage(new Uri("pack://application:,,,/CPI Sheets Updater;component/Resources/toggle-on-16.png"));
            button.ItemText = "Остановить" + Environment.NewLine + "Sheets Updater";
            button.ToolTip = "Следит за обновлением параметра Номер листа и заполняет параметры для ведомости чертежей и ведомости спецификаций для смежных листов (с теми же значениями параметров CPI_Группирование видов и ADSK_Назначение вида)";
        }

        public static void buttonOff() {
            button.LargeImage = new BitmapImage(new Uri("pack://application:,,,/CPI Sheets Updater;component/Resources/toggle-off.png"));
            button.Image = new BitmapImage(new Uri("pack://application:,,,/CPI Sheets Updater;component/Resources/toggle-off-16.png"));
            button.ItemText = "Запустить" + Environment.NewLine + "Sheets Updater";
            button.ToolTip = "Следит за обновлением параметра Номер листа и заполняет параметры для ведомости чертежей и ведомости спецификаций для смежных листов (с теми же значениями параметров CPI_Группирование видов и ADSK_Назначение вида)";
        }

        public static void buttonDead() {
            button.LargeImage = new BitmapImage(new Uri("pack://application:,,,/CPI Sheets Updater;component/Resources/toggle-dead.png"));
            button.Image = new BitmapImage(new Uri("pack://application:,,,/CPI Sheets Updater;component/Resources/toggle-dead-16.png"));
            button.ItemText = "Перезапустить" + Environment.NewLine + "Sheets Updater";
            button.ToolTip = "Параметры не найдены";
        }

        public Result OnShutdown(UIControlledApplication application) {
            SheetUpdater updater = new SheetUpdater(application.ActiveAddInId);
            UpdaterRegistry.UnregisterUpdater(updater.GetUpdaterId());
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application) {
            AddRibbonPanel(application);
            RegisterUpdater(application);
            application.ViewActivated += new EventHandler<Autodesk.Revit.UI.Events.ViewActivatedEventArgs>(App_ViewActivated);
            return Result.Succeeded;
        }

        void App_ViewActivated(object sender, Autodesk.Revit.UI.Events.ViewActivatedEventArgs e) {
            var doc = e.CurrentActiveView.Document;
            if (!UpdaterIsAvailable.ContainsKey(doc)) { UpdaterIsAvailable.Add(doc, CheckDoc(doc)); }
            //if (UpdaterIsAvailableInDoc[doc]) { return; }
            var flag = UpdaterIsAvailable[doc];
            if (flag) {
                if (UpdaterIsEnabledFlag) {
                    buttonOn();
                } else {
                    buttonOff();
                }
            } else {
                buttonDead();
            }
        }

        public static bool CheckDoc(Document doc) {
            foreach (var name in paramNames) {
                if (!SharedParameterIsBinded(doc, name, BuiltInCategory.OST_Sheets)) {
                    //s += $"\nНе найден параметр {name}";
                    return false;
                }
            }
            return true;
        }
        private static bool SharedParameterIsBinded(Document doc, string paraName, BuiltInCategory boundCategory) {
            try {
                BindingMap bindingMap = doc.ParameterBindings;
                DefinitionBindingMapIterator bindingMapIter = bindingMap.ForwardIterator();
                while (bindingMapIter.MoveNext()) {
                    if (bindingMapIter.Key.Name.Equals(paraName)) {
                        ElementBinding binding = bindingMapIter.Current as ElementBinding;
                        CategorySet categories = binding.Categories;
                        foreach (Category category in categories) {
                            if (category.Id.IntegerValue.Equals((int)boundCategory)) { return true; }
                        }
                    }
                }
            } catch (Exception) {
                return false;
            }

            return false;
        }

        private void RegisterUpdater(UIControlledApplication application) {
            // Register updater with Revit
            SheetUpdater updater = new SheetUpdater(application.ActiveAddInId);
            UpdaterRegistry.RegisterUpdater(updater);

            // Change Scope = any ViewSheet element
            ElementClassFilter SheetFilter = new ElementClassFilter(typeof(ViewSheet));

            // Change type = parameter + addition
            ElementId paramId = new ElementId(BuiltInParameter.SHEET_NUMBER);
            UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), SheetFilter, Element.GetChangeTypeParameter(paramId));
            UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), SheetFilter, Element.GetChangeTypeElementAddition());
        }
    }
    public class SheetUpdater : IUpdater {
        static AddInId m_appId;
        static UpdaterId m_updaterId;
        private Dictionary<ElementId, List<ScheduleSheetInstance>> _ssisByOwner;
        public SheetUpdater(AddInId id) {
            m_appId = id;
            m_updaterId = new UpdaterId(m_appId, new Guid("c99a8414-1b2d-41e5-980b-1a7115bcfb7c"));
        }

        public void Execute(UpdaterData data) {
            if (!App.UpdaterIsEnabledFlag) { return; }
            try {
                Document doc = data.GetDocument();
                var modSheets = new List<ViewSheet>();
                foreach (var id in data.GetModifiedElementIds().Concat(data.GetAddedElementIds())) {
                    Element elem = doc.GetElement(id);
                    if (elem is ViewSheet) {
                        modSheets.Add(elem as ViewSheet);
                    }
                }

                // check params are exist
                if (null == modSheets.First().LookupParameter("CPI_Количество листов")) { return; }

                //creating a dict {Owner: ScheduleSheetInstances}
                var SSInstances = new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance)).WhereElementIsNotElementType().ToElements();
                _ssisByOwner = new Dictionary<ElementId, List<ScheduleSheetInstance>>();
                foreach (ScheduleSheetInstance ssi in SSInstances) {
                    ElementId ownerId = ssi.OwnerViewId;
                    if (!_ssisByOwner.ContainsKey(ownerId)) {
                        _ssisByOwner.Add(ownerId, new List<ScheduleSheetInstance>());
                    }
                    _ssisByOwner[ownerId].Add(ssi);
                }
                //creating a dict {specName: sheets}
                var sheetsBySpecName = new Dictionary<string, List<ViewSheet>>();
                foreach (KeyValuePair<ElementId, List<ScheduleSheetInstance>> pair in _ssisByOwner) {
                    var ownerSheet = doc.GetElement(pair.Key) as ViewSheet;
                    var ssis = pair.Value;
                    var lst = new List<string>();
                    foreach (var ssi in ssis) {
                        if (1 == doc.GetElement(ssi.ScheduleId).LookupParameter("CPI_Исключить из ВС")?.AsInteger()) { continue; }
                        var text = (doc.GetElement(ssi.ScheduleId) as ViewSchedule).GetTableData().GetSectionData(SectionType.Header).GetCellText(0, 0);
                        if (text.Length > 0) {
                            if (text.Contains("Экспликация помещений")) { continue; }
                            lst.Add(text);
                        } else {
                            if (ssi.Name.Contains("Экспликация помещений")) { continue; }
                            List<string> arr = ssi.Name.Split().ToList();
                            if (arr.GetRange(arr.Count - 1, 1)[0].Contains("/")) { //if (revitVersion >= 2022 && ssi.SegmentIndex >= 0)
                                lst.Add(String.Join(" ", arr.GetRange(0, arr.Count - 1)));
                            } else {
                                lst.Add(ssi.Name);
                            }
                        }
                    }
                    HashSet<String> uniques = new HashSet<String>(lst);
                    lst = uniques.ToList();
                    lst.Sort();
                    var specName = String.Join("\n", lst);
                    ownerSheet.LookupParameter("CPI_ВС Наименование спецификации").Set(specName);

                    if (specName == "") { continue; }

                    if (!sheetsBySpecName.ContainsKey(specName)) {
                        sheetsBySpecName.Add(specName, new List<ViewSheet>());
                    }
                    sheetsBySpecName[specName].Add(ownerSheet);
                }

                #region 
                //creating a set of the combinations of the grouping parameters of the sheets, which are in the modyfied data
                HashSet<String> combinations = new HashSet<String>();
                foreach (ViewSheet sheet in modSheets) {
                    //ViewSheet sheet = doc.GetElement(elId) as ViewSheet;
                    String group = sheet.LookupParameter("CPI_Группирование видов")?.AsString() ?? "-";
                    String purpose = sheet.LookupParameter("ADSK_Назначение вида")?.AsString() ?? "-";
                    String group_purpose = group + purpose;
                    combinations.Add(group_purpose);
                }

                //filtering the sheets from the affected group
                FilteredElementCollector filter = new FilteredElementCollector(doc);
                var viewSheets = filter.OfClass(typeof(ViewSheet)).WhereElementIsNotElementType().ToElements();
                var sheets = (from sheet in viewSheets
                              where combinations.Contains((sheet.LookupParameter("CPI_Группирование видов")?.AsString() ?? "-")
                                                        + (sheet.LookupParameter("ADSK_Назначение вида")?.AsString() ?? "-"))
                              select sheet).ToList();

                if (doc.IsWorkshared) {
                    string username = doc.Application.Username;
                    var occupiedSheets = new Dictionary<string, List<ViewSheet>>();
                    foreach (ViewSheet sheet in sheets) {
                        string editor = sheet.get_Parameter(BuiltInParameter.EDITED_BY).AsString();
                        if (editor == "" || editor == username) { continue; }
                        if (!occupiedSheets.ContainsKey(editor)) { occupiedSheets.Add(editor, new List<ViewSheet>()); }
                        occupiedSheets[editor].Add(sheet);
                    }
                    if (occupiedSheets.Count() > 0) {
                        var s = "Листы, занятые другими пользователями, не будут обновлены.\n"
                            + string.Join("\n", occupiedSheets.Select(p => {
                                return $"Пользователь '{p.Key}' владеет листами:" +
                                $"{string.Join("", p.Value.Select(el => $"\n    {el.SheetNumber}"))}";
                            }).ToArray());
                        MessageBox.Show(s, "Предупреждение CPI Sheets Updater");
                        foreach (var sheet in occupiedSheets.Values.SelectMany(x => x).ToList()) {
                            sheets.Remove(sheet);
                        }
                    }
                }

                //creating a dict {intPart: sheets}
                var sheetsByInt = new Dictionary<int, List<ViewSheet>>();
                foreach (ViewSheet sheet in sheets) {
                    var intPart = Int32.Parse(GetParts(sheet).intPart);
                    if (!sheetsByInt.ContainsKey(intPart)) {
                        sheetsByInt.Add(intPart, new List<ViewSheet>());
                    }
                    sheetsByInt[intPart].Add(sheet);

                }

                #endregion
                // writing params of the coSheets
                foreach (KeyValuePair<int, List<ViewSheet>> pair in sheetsByInt) {
                    var coSheets = pair.Value;
                    var fracPartOfTheLastSheet = GetParts(coSheets.Last()).fracPart;
                    var coSheetsCount = coSheets.Count();
                    foreach (var sheet in coSheets) {
                        var parts = GetParts(sheet);
                        String intPart = parts.intPart;
                        String fracPart = parts.fracPart;
                        sheet.LookupParameter("CPI_Номер листа").Set(fracPart != "" ? $"{intPart}.{fracPart}" : $"{intPart}");
                        sheet.LookupParameter("CPI_Количество листов").Set(coSheets.Count == 1 ? "" : fracPartOfTheLastSheet);
                        var CPI_VCH_Nomer_lista = Compose_CPI_Nomer_lista(sheet, "1", fracPartOfTheLastSheet);
                        var num = sheet.LookupParameter("CPI_Номер листа").AsString();
                        sheet.LookupParameter("CPI_ВЧ Номер листа").Set(coSheets.Count == 1 ? num : CPI_VCH_Nomer_lista);
                        var prim = coSheets.Count == 1 ? "" : $"На {coSheetsCount} лист{GetEnding(coSheetsCount)}";
                        sheet.LookupParameter("CPI_ВЧ Примечание").Set(prim);

                        var specName = sheet.LookupParameter("CPI_ВС Наименование спецификации").AsString();
                        if (!_ssisByOwner.ContainsKey(sheet.Id))
                        {
                            specName = "";
                        }
                        if (null == specName) { continue; }
                        if ("" == specName) {
                            sheet.LookupParameter("CPI_ВС Номер листа").Set("");
                            sheet.LookupParameter("CPI_ВС Примечание").Set("");
                            sheet.LookupParameter("CPI_ВС Наименование спецификации").Set("");
                            continue;
                        }
                        if (sheetsBySpecName.ContainsKey(specName)) {
                            var coSheetsForVS = sheetsBySpecName[specName];
                            var fracPartOfTheLastSheetForVS = GetParts(coSheetsForVS.Last()).fracPart;
                            var fracPartOfTheFirstSheetForVS = GetParts(coSheetsForVS.First()).fracPart;
                            var coSheetsForVSCount = coSheetsForVS.Count();
                            var CPI_VS_Nomer_lista = Compose_CPI_Nomer_lista(sheet, fracPartOfTheFirstSheetForVS, fracPartOfTheLastSheetForVS);
                            sheet.LookupParameter("CPI_ВС Номер листа").Set(coSheetsForVS.Count == 1 ? num : CPI_VS_Nomer_lista);
                            prim = coSheetsForVS.Count == 1 ? "" : $"На {coSheetsForVSCount} лист{GetEnding(coSheetsForVSCount)}";
                            sheet.LookupParameter("CPI_ВС Примечание").Set(prim);
                        } else {
                            sheet.LookupParameter("CPI_ВС Номер листа").Set(num);
                        }
                    }
                }
            } catch (Exception ex) {
                var app = data.GetDocument().Application;
                var uiApp = new UIApplication(app);
                App.UpdaterIsEnabledFlag = false;
                App.buttonDead();
                App.button.ToolTip = "Непредвиденная ошибка";

                TaskDialog mainDialog = new TaskDialog("Ошибка");
                mainDialog.MainInstruction = "Средство обновления CPI Sheets Updater остановлено";
                mainDialog.ExpandedContent = ex.ToString();
                mainDialog.MainContent = "В средстве обновления CPI Sheets Updater возникла ошибка и оно было остановлено.\nДля запуска нажмите Надстройки, Запустить Sheets Updater";
                mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                                          "Запустить сейчас");
                mainDialog.CommonButtons = TaskDialogCommonButtons.Close;
                mainDialog.DefaultButton = TaskDialogResult.Close;
                TaskDialogResult tResult = mainDialog.Show();
                if (TaskDialogResult.CommandLink1 == tResult) {
                    App.buttonOn();
                    App.UpdaterIsEnabledFlag = true;
                }
            }
        }

        private string GetEnding(int count) {
            var ending = (singular: "е", plural: "ах");
            return count % 10 == 1 && count != 11 ? ending.singular : ending.plural; // "12.1 ÷ 12.42"
        }

        private string Compose_CPI_Nomer_lista(ViewSheet sheet, string fracPartOfTheFirstSheet, string fracPartOfTheLastSheet) {
            (string intPart, string _) = GetParts(sheet);
            return $"{intPart}.{fracPartOfTheFirstSheet} ÷ {intPart}.{fracPartOfTheLastSheet}"; // "12.1 ÷ 12.42"
        }

        public (string intPart, string fracPart) GetParts(ViewSheet sheet) {
            String numStr = sheet.LookupParameter("Номер листа").AsString();
            var separators = new char[] { ' ', '-', '_', '/' };
            String lastPart = numStr.Split(separators).Last();
            String intPart;
            String fracPart = "";
            if (lastPart.Contains(".")) {
                intPart = lastPart.Split('.')[0];
                fracPart = lastPart.Split('.')[1];
            } else {
                intPart = lastPart;
            }
            intPart = Regex.Replace(intPart, @"[^\d]", string.Empty);
            fracPart = Regex.Replace(fracPart, @"[^\d]", string.Empty);
            return (intPart, fracPart);
        }

        public string GetAdditionalInformation() {
            return "CPI Sheets Updater (Синхронизация номера листа): Синхронизирует значения вспомогательных параметров при создании и редактировании листа";
        }

        public ChangePriority GetChangePriority() {
            return ChangePriority.Views;
        }

        public UpdaterId GetUpdaterId() {
            return m_updaterId;
        }

        public string GetUpdaterName() {
            return "CPI Sheets Updater v0.1.2";
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class SwitchUpdater : IExternalCommand {

        public static RibbonItem GetRibbonItemByName(UIApplication app, String panelName, String itemName) {
            RibbonPanel panelRibbon = null;
            foreach (RibbonPanel item in app.GetRibbonPanels()) { if (panelName == item.Name) { panelRibbon = item; } }
            foreach (RibbonItem item in panelRibbon.GetItems()) { if (itemName == item.Name) { return item; } }
            return null;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {
            try {
                var doc = commandData.Application.ActiveUIDocument.Document;
                if (App.UpdaterIsAvailable[doc]) {
                    if (App.UpdaterIsEnabledFlag) {  // change state from Enable to Disable
                        App.buttonOff();
                        App.UpdaterIsEnabledFlag = false;
                    } else {
                        App.buttonOn();
                        App.UpdaterIsEnabledFlag = true;
                    }
                } else {
                    if (App.CheckDoc(doc)) {
                        App.UpdaterIsAvailable[doc] = true;
                        if (App.UpdaterIsEnabledFlag) {  // updater already is Enabled
                            App.buttonOn();
                        } else {
                            App.buttonOff();
                        }
                    } else {
                        App.buttonDead();
                    }
                }
                return Result.Succeeded;
            } catch (Exception ex) {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
