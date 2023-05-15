using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RLP
{
    [Transaction(TransactionMode.Manual)]
    public class Views : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            doc = uiDoc.Document;

            TaskDialogResult result = TaskDialog.Show("Выбор", "Выберите сборку", TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);

            if (result == TaskDialogResult.Cancel || result == TaskDialogResult.No)
            {
                return Result.Cancelled;
            }

            // Get an assembly instance from the user
            Reference assemblyRef = uiDoc.Selection.PickObject(ObjectType.Element, new AssemblySelectionFilter(), "Выберите сборку");
            try
            {
                AssemblyInstance assemblyInstance = doc.GetElement(assemblyRef) as AssemblyInstance;

                if (assemblyInstance != null)
                {
                    string assebmlyName = assemblyInstance.Name;
                    string viewName = $"Наружная стеновая панель {assebmlyName}. Опалубка";
                    bool newViews = true;

                    // Check if the view already exists
                    bool viewExists = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Any(x => x.Name == viewName);

                    // Ask the user if they want to create the view anyway
                    if (viewExists)
                    {
                        TaskDialogResult overwriteResult = TaskDialog.Show("Вид уже существует", $"Вид {viewName} уже существует. Хотите создать новый вид с другим именем ({viewName} (1))?", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                        if (overwriteResult == TaskDialogResult.No)
                        {
                            newViews = false;
                        }
                        else
                        {
                            // Append a number to the view name
                            int viewNumber = 1;
                            while (new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Any(x => x.Name == $"{viewName} ({viewNumber})"))
                            {
                                viewNumber++;
                            }
                            viewName = $"{viewName} ({viewNumber})";
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "Создать опалубку"))
                    {
                        trans.Start();
                        
                        /////Creation of oopalubka

                        ViewSheet sheet = CreateNewSheet(assebmlyName);
                        ViewSection sectionViewBack = null;

                        if (newViews)
                        {
                            // Create a section view of the assembly instance type elevation back
                            sectionViewBack = AssemblyViewUtils.CreateDetailSection(doc, assemblyInstance.Id, AssemblyDetailViewOrientation.ElevationBack);
                            sectionViewBack.Name = viewName;
                            ApplyViewTemplate(sectionViewBack, "3НСНг_Панель_Фасад_Опалубка");
                            sectionViewBack.get_Parameter(BuiltInParameter.VIEW_NAME).Set(viewName);
                            sectionViewBack.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(assebmlyName);
                            // Get the view you want to modify
                            View viewToModify = sectionViewBack;

                            // Get all elements in the view that have the same name as the view
                            List<ElementId> elementsToHide = new FilteredElementCollector(doc)
                                .OfClass(typeof(View))
                                .Where(v => v.Name == viewName)
                                .Select(v => v.Id)
                                .ToList();

                            try { viewToModify.HideElements(elementsToHide); } catch { elementsToHide.Clear(); elementsToHide.Add(viewToModify.Id); }
                            try { viewToModify.HideElements(elementsToHide); } catch { }
                            try { viewToModify.HideElements((ICollection<ElementId>)viewToModify.Id); } catch { }
                        }
                        else 
                        {
                            sectionViewBack = (ViewSection)new FilteredElementCollector(doc)
                            .OfClass(typeof(View))
                            .Cast<View>()
                            .FirstOrDefault(view => view.Name.Equals(viewName));

                            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet));

                            // Loop through each sheet
                            foreach (ViewSheet anysheet in sheetCollector)
                            {
                                // Check if the view is present in any of the viewports
                                foreach (ElementId viewportId in anysheet.GetAllViewports())
                                {
                                    Viewport viewport1 = (Viewport)doc.GetElement(viewportId);
                                    if (viewport1.ViewId == sectionViewBack.Id)
                                    {
                                        // The view is already on a sheet
                                        TaskDialog.Show("Вид на листе", $"Вид {sectionViewBack.Name} уже размещен на листе. Название листа: " + sheet.Name);
                                        return Result.Failed;
                                    }
                                }
                            }
                        }

                        Viewport viewport = Viewport.Create(doc, sheet.Id, sectionViewBack.Id, new XYZ(0, 0, 0 ));
                        //viewport.LabelOffset = XYZ.Zero;

                        //TaskDialogResult overwriteResult = TaskDialog.Show("Выровнять вид", $"Выровнять вид на листе?", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                        //if (overwriteResult == TaskDialogResult.No)
                        //{
                        //    newViews = false;
                        //}
                        //else
                        //{
                            // Get the outline of the sheet
                            //var outline = sheet.Outline;
                            var viewOutline = viewport.GetBoxOutline();


                            // get the sheet outline and viewport bounding box
                            BoundingBoxUV sheetBox = sheet.Outline;
                            BoundingBoxXYZ viewportBox = viewport.get_BoundingBox(sheet);

                            // calculate the center points
                            UV sheetCenter = (sheetBox.Max + sheetBox.Min) / 2;
                            XYZ viewportCenter = (viewportBox.Max + viewportBox.Min) / 2;

                            // calculate the offset to center the viewport on the sheet
                            double offsetX = sheetCenter.U - viewportCenter.X;
                            double offsetY = sheetCenter.V - viewportCenter.Y;

                            // set the viewport position
                            viewport.SetBoxCenter(new XYZ(viewportBox.Max.X + offsetX * 2.5, viewportBox.Max.Y - offsetY, 0));

                        //}

                        // set the label offset to top and center
                        //XYZ labelOffset = new XYZ(viewportBox.Max.X + offsetX * 2.5, viewportBox.Max.Y * 2 - offsetY, 0);
                        //viewport.LabelOffset = labelOffset;

                        SetViewportType(sectionViewBack, "Заголовок на листе");

                        trans.Commit();

                        try
                        {
                            // Show the newly created views
                            uiDoc.ActiveView = sectionViewBack;
                            uiDoc.ActiveView = sheet;
                            var ports = sheet.GetAllViewports();

                        }
                        catch (Exception ex)
                        {
                            TaskDialog taskDialog = new TaskDialog("Something went wrong...");
                            taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                            taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                            taskDialog.Show();
                        }
                    }




                    return Result.Succeeded;
                }
                else
                {
                    message = "Не удалось получить экземпляр сборки.";
                    return Result.Failed;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // The user aborted the PickObject operation, so cancel the command
                return Result.Cancelled;
            }
            catch (Exception ex) {
                TaskDialog taskDialog = new TaskDialog("Something went wrong... (The whole views command)");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
                return Result.Failed;
            }
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

        private void ApplyViewTemplate(View view, string viewTemplateName)
        {
            // Get the view template element
            View viewTemplate = new FilteredElementCollector(view.Document)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

            if (viewTemplate == null)
            {
                TaskDialog.Show("Error", $"The view template named {viewTemplateName} does not exist.");
                return;
            }

            try
            {
                // Set the view template parameter of the view
                using (Transaction tx = new Transaction(view.Document, "Apply View Template"))
                {
                    tx.Start();
                    view.ViewTemplateId = viewTemplate.Id;
                    tx.Commit();
                }
            }
            catch (Exception ex) { view.ViewTemplateId = viewTemplate.Id; }
        }

        public static List<Viewport> GetViewports(View view)
        {
            var collector = new FilteredElementCollector(view.Document)
                .OfClass(typeof(Viewport))
                .WhereElementIsNotElementType()
                .Cast<Viewport>();

            var viewports = collector.Where(vp => vp.ViewId == view.Id).ToList();
            return viewports;
        }

        public void SetViewportType(View view, string viewportTypeName)
        {
            try
            {
                foreach (Viewport viewport in GetViewports(view))
                { 
                    ElementType scaledViewPortType = null;
                    ElementType currentViewprtType = doc.GetElement(viewport.GetTypeId()) as ElementType;

                    if (currentViewprtType.Name != viewportTypeName)
                    {
                        foreach (ElementId id in viewport.GetValidTypes())
                        {
                            ;
                            ElementType type = doc.GetElement(id) as ElementType;
                            //TaskDialog.Show("Tak", type.Name);
                            if (type.Name == viewportTypeName)
                            {
                                scaledViewPortType = type;
                                break;
                            }
                        }
                    }
                    if (scaledViewPortType != null)
                    {
                        try
                        {
                            using (Transaction t = new Transaction(doc, "Смена типа"))
                            {
                                t.Start();
                                viewport.ChangeTypeId(scaledViewPortType.Id);
                                doc.Regenerate();
                                t.Commit();
                            }
                        }
                        catch
                        {
                            viewport.ChangeTypeId(scaledViewPortType.Id);
                            doc.Regenerate();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Something went wrong... (Set viewport)");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
            }
        }
        public ViewSheet CreateNewSheet(string assemblyName)
        {
            // Get the title block type id for the desired type
            FilteredElementCollector titleBlocks = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks);
            FamilySymbol titleBlock = titleBlocks
                .Where(x => x.Name == "Форма 4_Панели (масштаб 1 к 20)").FirstOrDefault() as FamilySymbol;
            ElementId titleBlockTypeId = titleBlock.Id;

            // Create the new sheet
            ViewSheet newSheet = ViewSheet.Create(doc, titleBlockTypeId);
            newSheet.Name = $"Наружная стеновая панель {assemblyName}";
            return newSheet;
        }
    }
}

