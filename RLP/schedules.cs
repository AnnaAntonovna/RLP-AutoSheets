using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Structure;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Reflection.Emit;
using System.Windows;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Windows.Shapes;
using Line = Autodesk.Revit.DB.Line;
using System.Xml.Linq;
using System.Drawing;
using System.Windows.Media.Media3D;
using System.Windows.Controls;
using System.Windows.Forms;
using View = Autodesk.Revit.DB.View;
using Autodesk.Revit.Creation;

namespace RLP
{
    [Transaction(TransactionMode.Manual)]
    public class ScheduleCMD : IExternalCommand
    {
        Autodesk.Revit.DB.Document doc { get; set; }
        UIDocument uiDoc { get; set; }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiDoc = commandData.Application.ActiveUIDocument;
            doc = uiDoc.Document;
            ElementId categoryId = Category.GetCategory(doc, BuiltInCategory.OST_Rebar).Id;

            TaskDialog taskDialog = new TaskDialog("Выбор сборок");
            taskDialog.MainInstruction = "Выбор сборок";
            taskDialog.MainContent = "Метод выбора:";
            taskDialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Add command links with links
            taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Выбор вручную");
            taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Автоматический выбор внешних панелей 3НСН");
            taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Автоматичекий выбор внешних панелей 1НСН");
            taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Автоматичекий выбор внутренних панелей ПСВ");

            List<AssemblyInstance> assemblies = new List<AssemblyInstance>();

            TaskDialogResult tResult = taskDialog.Show();

            if (tResult == TaskDialogResult.CommandLink1) // Manual selection
            {
                TaskDialogResult resulte = TaskDialog.Show("Выбор", "Выберите сборку", TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);

                if (resulte == TaskDialogResult.Cancel || resulte == TaskDialogResult.No)
                {
                    return Result.Cancelled;
                }

                // Get an assembly instance from the user
                Reference assemblyRef1 = uiDoc.Selection.PickObject(ObjectType.Element, new AssemblySelectionFilter(), "Выберите сборку");
                AssemblyInstance assemblyInstance1 = doc.GetElement(assemblyRef1) as AssemblyInstance;
                assemblies.Add(assemblyInstance1);

            }
            else if (tResult == TaskDialogResult.CommandLink2)
            {
                assemblies = new FilteredElementCollector(doc)
                    .OfClass(typeof(AssemblyInstance))
                    .Where(a => a.Name.Contains("3НСН"))
                    .Cast<AssemblyInstance>()
                    .ToList();
            }
            else if (tResult == TaskDialogResult.CommandLink3)
            {
                assemblies = new FilteredElementCollector(doc)
                    .OfClass(typeof(AssemblyInstance))
                    .Where(a => a.Name.Contains("1НСН"))
                    .Cast<AssemblyInstance>()
                    .ToList();
            }
            else if (tResult == TaskDialogResult.CommandLink4)
            {
                assemblies = new FilteredElementCollector(doc)
                    .OfClass(typeof(AssemblyInstance))
                    .Where(a => a.Name.Contains("ПСВ"))
                    .Cast<AssemblyInstance>()
                    .ToList();
            }

            foreach (AssemblyInstance assemblyInstance in assemblies)
            {
                try
                {
                    if (assemblyInstance != null)
                    {
                        string assebmlyName = assemblyInstance.Name;
                        if (assebmlyName.Contains("3НСН"))
                        {
                            double height = 0;
                            ViewSchedule schedule1 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 1.1_Каркасы_3НСНг, 1НСН, 1НС", "Спец № 1_Каркасы123", categoryId); 
                            ViewSchedule schedule2 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 2_Сетки", "Спец № 2_Сетки123", categoryId);
                            ViewSchedule schedule3 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 3_Закладные детали", "Спец № 3_Закладные123", categoryId);
                            ViewSchedule schedule4 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 4_Стержни и гнутые детали", "Спец № 4_Арматура123", categoryId);
                            ViewSchedule schedule5 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 5_Стандартные изделия", "Спец № 5_Стандартные изделия123", categoryId);

                            ViewSchedule schedule7 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "3НСНг_Ведомость расхода стали на изделия арматурные", "ВРС_Арматура123", categoryId);
                            ViewSchedule schedule8 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "3НСНг_Ведомость расхода стали на изделия закладные", "ВРС_Закладные123", categoryId);


                            ViewSheet sheet = CreateSheetWithAssembly(doc, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Спецификации"
                                , 2, "5", "Форма 6");

                            PlaceScheduleOnSheet(sheet);

                            using (Transaction transaction = new Transaction(doc, "Place schedules on a sheet"))
                            {
                                try
                                {
                                    transaction.Start();
                                    // Create the viewport on the sheet
                                    try
                                    {
                                        Viewport viewport1 = Viewport.Create(sheet.Document, sheet.Id, schedule1.Id, XYZ.Zero);
                                        height = height + GetViewportHeight(sheet, viewport1);
                                        ChangeViewportPlaceOnSheet(viewport1, schedule1, sheet, 0.2, 0.9 - height);
                                    }
                                    catch (Exception ex) { ShowException("Could not place schedules on a sheet", ex); }

                                    try
                                    {
                                        Viewport viewport2 = Viewport.Create(sheet.Document, sheet.Id, schedule2.Id, XYZ.Zero);
                                    height = height + GetViewportHeight(sheet, viewport2);
                                    ChangeViewportPlaceOnSheet(viewport2, schedule2, sheet, 0.2, 0.9 - height);
                                    }
                                    catch (Exception ex) { ShowException("Could not place schedules on a sheet", ex); }

                                    try
                                    {
                                        Viewport viewport3 = Viewport.Create(sheet.Document, sheet.Id, schedule3.Id, XYZ.Zero);
                                    height = height + GetViewportHeight(sheet, viewport3);
                                    ChangeViewportPlaceOnSheet(viewport3, schedule3, sheet, 0.2, 0.9 - height);
                                    }
                                    catch (Exception ex) { ShowException("Could not place schedules on a sheet", ex); }

                                    try
                                    {
                                        Viewport viewport4 = Viewport.Create(sheet.Document, sheet.Id, schedule4.Id, XYZ.Zero);
                                    height = height + GetViewportHeight(sheet, viewport4);
                                    ChangeViewportPlaceOnSheet(viewport4, schedule4, sheet, 0.2, 0.9 - height);
                                    }
                                    catch (Exception ex) { ShowException("Could not place schedules on a sheet", ex); }

                                    try
                                    {
                                        Viewport viewport5 = Viewport.Create(sheet.Document, sheet.Id, schedule5.Id, XYZ.Zero);
                                    height = height + GetViewportHeight(sheet, viewport5);
                                    ChangeViewportPlaceOnSheet(viewport5, schedule5, sheet, 0.2, 0.9 - height);
                                    }
                                    catch (Exception ex) { ShowException("Could not place schedules on a sheet", ex); }

                                    transaction.Commit();
                                }
                                catch (Exception ex) { ShowException("Could not place schedules on a sheet", ex); }
                                // Change the placement of the viewport on the sheet
                            }

                            uiDoc.ActiveView = sheet;
                        }
                        else if (assebmlyName.Contains("1НСН"))
                        {
                            ViewSchedule schedule1 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 1.1_Каркасы_3НСНг, 1НСН, 1НС", "Спец № 1_Каркасы", categoryId);
                            ViewSchedule schedule2 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 2_Сетки", "Спец № 2_Сетки", categoryId);
                            ViewSchedule schedule3 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 3_Закладные детали", "Спец № 3_Закладные", categoryId);
                            ViewSchedule schedule4 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 4_Стержни и гнутые детали", "Спец № 4_Арматура", categoryId);
                            ViewSchedule schedule5 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 5_Стандартные изделия", "Спец № 5_Стандартные изделия", categoryId);

                            ViewSchedule schedule7 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "НСН-ПСВ_Ведомость расхода стали на изделия арматурные", "ВРС_Арматура", categoryId);
                            ViewSchedule schedule8 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "НСН-ПСВ_Ведомость расхода стали на изделия закладные", "ВРС_Закладные", categoryId);


                            ViewSheet sheet = CreateSheetWithAssembly(doc, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Спецификации"
                                , 2, "3", "Форма 6");
                            uiDoc.ActiveView = sheet;
                        }
                        else if (assebmlyName.Contains("ПСВ"))
                        {
                            ViewSchedule schedule1 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "Спецификация к панелям № 1.2_Каркасы_ПСВ", "Спец № 1_Каркасы", categoryId);
                            ViewSchedule schedule2 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 2_Сетки", "Спец № 2_Сетки", categoryId);
                            ViewSchedule schedule3 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 3_Закладные детали", "Спец № 3_Закладные", categoryId);
                            ViewSchedule schedule4 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 4_Стержни и гнутые детали", "Спец № 4_Арматура", categoryId);
                            ViewSchedule schedule5 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                "Спецификация к панелям № 5_Стандартные изделия", "Спец № 5_Стандартные изделия", categoryId);

                            ViewSchedule schedule7 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "НСН-ПСВ_Ведомость расхода стали на изделия арматурные", "ВРС_Арматура", categoryId);
                            ViewSchedule schedule8 = CreateScheduleFromTemplate(doc, assemblyInstance,
                                    "НСН-ПСВ_Ведомость расхода стали на изделия закладные", "ВРС_Закладные", categoryId);

                            ViewSheet sheet = CreateSheetWithAssembly(doc, assemblyInstance, $"Внутренняя стеновая панель {assebmlyName}. Спецификации"
                                , 2, "3", "Форма 6");
                            uiDoc.ActiveView = sheet;
                        }
                        else
                        { TaskDialog.Show("Oh..", "Название сборки не корректное. Листы не созданы"); return Result.Failed; }
                    }
                }
                catch (Exception ex) { ShowException($"Could not do {assemblyInstance.Name}",ex);}
            }
            return Result.Succeeded;
        }

        public ViewSchedule CreateScheduleFromTemplate(Autodesk.Revit.DB.Document document, AssemblyInstance assembly, string templateName, string scheduleName, ElementId categoryId)
        {

            // Create a schedule based on the template
            using (Transaction transaction = new Transaction(document, "Create Schedule"))
            {
                try
                {
                    transaction.Start();

                    ElementId viewtemplateid = GetViewTemplateId(templateName);
                    // Create the schedule
                    ViewSchedule schedule = AssemblyViewUtils.CreateSingleCategorySchedule(doc, assembly.Id, categoryId, viewtemplateid, true);


                    if (schedule == null)
                    {
                        TaskDialog.Show("Error", "Failed to create schedule.");
                        transaction.RollBack();
                        return null;
                    }

                    // Set the schedule name
                    schedule.Name = scheduleName;

                    // Regenerate the document to update the schedule
                    document.Regenerate();

                    transaction.Commit();

                    uiDoc.ActiveView = schedule;

                    return schedule;
                }
                catch (Exception ex) { ShowException("Could not CreateScheduleFromTemplate", ex); return null; }
            }
        }

        private ElementId GetViewTemplateId(string viewTemplateName)
        {
            // Get the view template element
            View viewTemplate = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

            if (viewTemplate == null)
            {
                TaskDialog.Show("Error", $"Шаблон вида {viewTemplateName} не существует в проекте. (Или не создано ни единой спецификации с применением этого шаблона).");
                return null;
            }
            try
            {
                // Set the view template parameter of the view

                return viewTemplate.Id;

            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't GetViewTemplateId");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
            }
            return null;
        }
        public static void ShowException(string errorName, Exception ex)
        {
            TaskDialog taskDialog = new TaskDialog("Oh-oh, " + errorName);
            taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
            taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
            taskDialog.Show();
        }
        public ViewSheet CreateSheetWithAssembly(Autodesk.Revit.DB.Document document, Element assemblyInstance, string sheetName, 
            int aForm, string sheetNumber, string titleBlockName)
        {
            // Check if a sheet with the given name already exists
            ViewSheet existingSheet = (ViewSheet)new FilteredElementCollector(document)
                .OfClass(typeof(ViewSheet))
                .FirstOrDefault(x => x.Name.Equals(sheetName)&&((ViewSheet)x).IsAssemblyView);

            if (existingSheet != null)
            {
                return (ViewSheet)existingSheet;
            }

            // Create a new sheet
            Transaction transaction = new Transaction(document, "Create Sheet");
            ElementId assemblyInstanceId = assemblyInstance.Id;

            // Get the title block type id for the desired type
            FilteredElementCollector titleBlocks = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks);
            FamilySymbol titleBlock = titleBlocks
                .Where(x => x.Name == titleBlockName).FirstOrDefault() as FamilySymbol;
            ElementId titleBlockTypeId = titleBlock.Id;


            transaction.Start();

            // Create the assembly view
            ViewSheet sheet = AssemblyViewUtils.CreateSheet(document, assemblyInstanceId, titleBlockTypeId);

            // Create the sheet

            // Set the sheet number and name
            sheet.SheetNumber = sheetName;
            sheet.Name = sheetName;


            Element titleBlockInstance = new FilteredElementCollector(doc, sheet.Id)
                            .OfCategory(BuiltInCategory.OST_TitleBlocks)
                            .ToElements()
                            .ToList().First();

            //Set parameters ("Форма А") - aForm
            Parameter titleForm = titleBlockInstance.LookupParameter("Формат А");
            if (titleForm != null)
            {
                titleForm.Set(aForm);
            }
            else { TaskDialog.Show("Can't set", $"Форма А does not set to '{aForm}'"); }

            //SHEET
            //"RLP_Номер листа" - sheetNumber;
            Parameter nomerLista = sheet.LookupParameter("RLP_Номер листа");
            if (nomerLista != null)
            {
                nomerLista.Set(sheetNumber);
            }
            else { TaskDialog.Show("Can't set", $"RLP_Номер листа does not set"); }

            SetSheetParameterValues(sheet, $"КЖ1.И1-{assemblyInstance.Name}",
                    "Конструкции железобетонных и арматурных изделий. Надземная часть. Панели наружные",
                     5);

            transaction.Commit();

            return sheet;
        }
        private void SetSheetParameterValues(ViewSheet sheet, string razdel, string naimenovanie, int kolvo)
        {
            try
            {

                // Get the parameters by name
                Parameter parameter1 = sheet.LookupParameter("ADSK_Штамп_Раздел проекта");
                Parameter parameter2 = sheet.LookupParameter("ADSK_Штамп_Наименование объекта");
                Parameter parameter3 = sheet.LookupParameter("ADSK_Штамп_Количество листов");

                // Set the parameter values
                if (parameter1 != null)
                {
                    parameter1.Set(razdel);
                }

                if (parameter2 != null)
                {
                    parameter2.Set(naimenovanie);
                }

                if (parameter3 != null)
                {
                    parameter3.Set(kolvo);
                }
            }
            catch (Exception ex) { ShowException("Can't CreateNewSheet", ex); }
        }

        public void PlaceScheduleOnSheet(ViewSheet sheet)
        {
            string scheduleName = "Заголовок для спецификаций";

            // Get the schedule view by name
            ViewSchedule schedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(x => x.Name.Equals(scheduleName));

            if (schedule == null)
            {
                // Schedule not found, display an error message or handle it accordingly
                TaskDialog.Show("Error", $"Schedule '{scheduleName}' not found.");
                return;
            }
                // Create a new transaction
            using (Transaction transaction = new Transaction(doc, "Place Schedule"))
            {
                transaction.Start();
                try
                {
                    // Place the schedule on the view
                    Viewport viewport = Viewport.Create(sheet.Document, sheet.Id, schedule.Id, XYZ.Zero);
                    ChangeViewportPlaceOnSheet(viewport, schedule, sheet, 0.2, 0.9);
                }
                catch (Exception ex) { ShowException("SetSheetParameterValues", ex);  }
                transaction.Commit();
            }
            
        }
        public double GetViewportHeight(ViewSheet sheet, Viewport viewport)
        {
            // Get the bounding box of the viewport on the sheet
            var viewportBox = viewport.GetBoxOutline();

            var minPoint = viewportBox.MinimumPoint;
            var maxPoint = viewportBox.MaximumPoint;

            // Calculate and return the height of the viewport
            double viewportHeight = maxPoint.Y - minPoint.Y;
            return viewportHeight;
        }
        public void ChangeViewportPlaceOnSheet(Viewport viewport, View view, ViewSheet sheet, double x, double y)
        {
            BoundingBoxXYZ boundingBox = sheet.get_BoundingBox(null);
            XYZ minPoint = boundingBox.Min;
            XYZ maxPoint = boundingBox.Max;

            XYZ vectorLlegend1 = new XYZ((-sheet.Outline.Max.U + sheet.Outline.Min.U) * x, (+sheet.Outline.Max.V - sheet.Outline.Min.V) * y, 0);
            viewport.Location.Move(vectorLlegend1);
        }
        public Viewport GetViewportAndParameters(ViewSheet sheet, ViewSchedule view, double x, double y)
        {
            // Create the viewport on the sheet
            Viewport viewport = Viewport.Create(sheet.Document, sheet.Id, view.Id, XYZ.Zero);

            // Calculate the height of the viewport
            double viewportHeight = GetViewportHeight(sheet, viewport);

            // Change the placement of the viewport on the sheet
            ChangeViewportPlaceOnSheet(viewport, view, sheet, x, 0.9 - viewportHeight);

            // Return the created viewport
            return viewport;
        }

        public class AssemblySelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                return element is AssemblyInstance;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return true;
            }
        }
    } 
}
