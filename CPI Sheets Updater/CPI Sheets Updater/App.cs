using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;

namespace CPI_Sheets_Updater {
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication {
        static readonly string AddInPath = typeof(App).Assembly.Location;
        static readonly string ButtonIconsFolder = Path.GetDirectoryName(AddInPath);
        static void AddRibbonPanel(UIControlledApplication application) {
                        
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("53 ЦПИ");
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData pbData = new PushButtonData(
                "cmdDisableUpdater",
                "Остановить" + System.Environment.NewLine + "Sheets Updater",
                thisAssemblyPath,
                "CPI_Sheets_Updater.DisableUpdater");
            PushButton pb = ribbonPanel.AddItem(pbData) as PushButton;
            pb.ToolTip = "Тooltip is coming soon";
            pb.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "icon.png"), UriKind.Absolute));

            PushButtonData pbData2 = new PushButtonData(
                "cmdEnableUpdater",
                "Запустить" + System.Environment.NewLine + "Sheets Updater",
                thisAssemblyPath,
                "CPI_Sheets_Updater.EnableUpdater");
            PushButton pb2 = ribbonPanel.AddItem(pbData2) as PushButton;
            pb2.ToolTip = "Тooltip is coming soon";
            pb2.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "icon.png"), UriKind.Absolute));
        }
        public Result OnShutdown(UIControlledApplication application) {
            SheetUpdater updater = new SheetUpdater(application.ActiveAddInId);
            UpdaterRegistry.UnregisterUpdater(updater.GetUpdaterId());
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application) {
            AddRibbonPanel(application);
            RegisterUpdater(application);
            return Result.Succeeded;
        }

        private void RegisterUpdater(UIControlledApplication application) {
            // Register updater with Revit
            SheetUpdater updater = new SheetUpdater(application.ActiveAddInId);
            UpdaterRegistry.RegisterUpdater(updater);

            // Change Scope = any ViewSheet element
            ElementClassFilter SheetFilter = new ElementClassFilter(typeof(ViewSheet));
            //ElementClassFilter ViewScheduleFilter = new ElementClassFilter(typeof(ViewSchedule));

            // Change type = parameter + addition
            ElementId paramId = new ElementId(BuiltInParameter.SHEET_NUMBER);
            UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), SheetFilter, Element.GetChangeTypeParameter(paramId));
            UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), SheetFilter, Element.GetChangeTypeElementAddition());
        }
    }
    public class SheetUpdater : IUpdater {
        static AddInId m_appId;
        static UpdaterId m_updaterId;
        public static bool UpdaterIsEnabledFlag = true;
        public SheetUpdater(AddInId id) {
            m_appId = id;
            m_updaterId = new UpdaterId(m_appId, new Guid("c99a8414-1b2d-41e5-980b-1a7115bcfb7c"));
        }

        public void Execute(UpdaterData data) {
            if (!UpdaterIsEnabledFlag) { return; }
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
                var ssisByOwner = new Dictionary<ElementId, List<ScheduleSheetInstance>>();
                foreach (ScheduleSheetInstance ssi in SSInstances) {
                    ElementId ownerId = ssi.OwnerViewId;
                    if (!ssisByOwner.ContainsKey(ownerId)) {
                        ssisByOwner.Add(ownerId, new List<ScheduleSheetInstance>());
                    }
                    ssisByOwner[ownerId].Add(ssi);
                }
                //creating a dict {specName: sheets}
                var sheetsBySpecName = new Dictionary<string, List<ViewSheet>>();
                foreach (KeyValuePair<ElementId, List<ScheduleSheetInstance>> pair in ssisByOwner) {
                    var ownerSheet = doc.GetElement(pair.Key) as ViewSheet;
                    var ssis = pair.Value;
                    var lst = new List<string>();
                    foreach (var ssi in ssis) {
                        var text = (doc.GetElement(ssi.ScheduleId) as ViewSchedule).GetTableData().GetSectionData(SectionType.Header).GetCellText(0, 0);
                        if (text.Length > 0) {
                            lst.Add(text);
                        } else {
                            lst.Add(ssi.Name);
                        }
                    }
                    var specName = String.Join("\n", lst);
                    ownerSheet.LookupParameter("CPI_ВС Наименование спецификации").Set(specName);

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
                    String group = sheet.LookupParameter("CPI_Группирование видов").AsString();
                    String purpose = sheet.LookupParameter("ADSK_Назначение вида").AsString();
                    String group_purpose = group + purpose;
                    combinations.Add(group_purpose);
                }

                //filtering the sheets from the affected group
                FilteredElementCollector filter = new FilteredElementCollector(doc);
                var viewSheets = filter.OfClass(typeof(ViewSheet)).WhereElementIsNotElementType().ToElements();
                var sheets = from sheet in viewSheets
                             where combinations.Contains(sheet.LookupParameter("CPI_Группирование видов").AsString() + sheet.LookupParameter("ADSK_Назначение вида").AsString())
                             select sheet;

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
                        if (null == specName) { continue; }
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
                UpdaterIsEnabledFlag = false;
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
                    UpdaterIsEnabledFlag = true;
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
            var numArr = numStr.Split(' ', '.');
            String intPart;
            String fracPart = "";
            if (numStr.Contains(" ") && numStr.Contains(".")) {
                intPart = numArr[1];
                fracPart = numArr[2];
            } else if (numStr.Contains(" ")) {
                intPart = numArr[1];
            } else if (numStr.Contains(".")) {
                intPart = numArr[0];
                fracPart = numArr[1];
            } else {
                intPart = numStr;
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
            return "CPI Sheets Updater v0.1.0";
        }
    }
    [Transaction(TransactionMode.Manual)]
    public class DisableUpdater : IExternalCommand {
        public Result Execute(
                ExternalCommandData commandData,
                ref string message,
                ElementSet elements) {

            try {
                SheetUpdater.UpdaterIsEnabledFlag = false;
                //TaskDialog.Show("DisableUpdater", "Средство обновления CPI Sheets Updater остановлено");
                return Result.Succeeded;
            } catch (Exception ex) {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
    [Transaction(TransactionMode.Manual)]
    public class EnableUpdater : IExternalCommand {
        public Result Execute(
                ExternalCommandData commandData,
                ref string message,
                ElementSet elements) {

            try {
                SheetUpdater.UpdaterIsEnabledFlag = true;
                //TaskDialog.Show("EnableUpdater", "Средство обновления CPI Sheets Updater запущено");
                return Result.Succeeded;
            } catch (Exception ex) {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
