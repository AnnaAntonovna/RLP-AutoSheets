using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
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

                    // Check if the view already exists
                    bool viewExists = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Any(x => x.Name == viewName);

                    // Ask the user if they want to create the view anyway
                    if (viewExists)
                    {
                        TaskDialogResult overwriteResult = TaskDialog.Show("View already exists", "A view with the same name already exists. Do you want to create a new view with a different name?", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                        if (overwriteResult == TaskDialogResult.No)
                        {
                            return Result.Cancelled;
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

                    using (Transaction trans = new Transaction(doc, "Создать виды сборки"))
                    {
                        trans.Start();
                        
                        ViewSheet sheet = CreateNewSheet(assebmlyName);


                        // Create a section view of the assembly instance Type A
                        ViewSection sectionView = AssemblyViewUtils.CreateDetailSection(doc, assemblyInstance.Id, AssemblyDetailViewOrientation.DetailSectionA);
                        sectionView.Name = "Section View of " + assemblyInstance.Name;
                        ApplyViewTemplate(sectionView, "3НСНг_Панель_Фасад_Опалубка");

                        // Create a section view of the assembly instance type B
                        ViewSection sectionViewB = AssemblyViewUtils.CreateDetailSection(doc, assemblyInstance.Id, AssemblyDetailViewOrientation.DetailSectionB);
                        sectionView.Name = "Section View of " + assemblyInstance.Name;
                        ApplyViewTemplate(sectionViewB, "3НСНг_Панель_Фасад_Опалубка");

                        // Create a section view of the assembly instance elevation front
                        ViewSection sectionViewFront = AssemblyViewUtils.CreateDetailSection(doc, assemblyInstance.Id, AssemblyDetailViewOrientation.ElevationFront);
                        sectionView.Name = "Section View of " + assemblyInstance.Name;
                        ApplyViewTemplate(sectionViewFront, "3НСНг_Панель_Фасад_Опалубка");

                        // Create a section view of the assembly instance type elevation back
                        ViewSection sectionViewBack = AssemblyViewUtils.CreateDetailSection(doc, assemblyInstance.Id, AssemblyDetailViewOrientation.ElevationBack);
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

                        try { viewToModify.HideElements(elementsToHide); } catch { }
                        try { viewToModify.HideElements((ICollection<ElementId>)viewToModify.Id); } catch { }


                        // Add the view to the sheet
                        BoundingBoxXYZ boundingBox = sheet.get_BoundingBox(null);
                        XYZ center = - (boundingBox.Max + boundingBox.Min) / 2;
                        Viewport viewport = Viewport.Create(doc, sheet.Id, sectionViewBack.Id, center);
                        
                        SetViewportType(sectionViewBack, "Заголовок на листе");

                        trans.Commit();
                        try
                        {
                            // Show the newly created views
                            uiDoc.ActiveView = sectionView;
                            uiDoc.ActiveView = sectionViewB;
                            uiDoc.ActiveView = sectionViewFront;
                            uiDoc.ActiveView = sectionViewBack;
                            uiDoc.ActiveView = sheet;
                            var ports = sheet.GetAllViewports();
                            TaskDialog.Show("ViewPorst", ports.First().ToString());


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
                            TaskDialog.Show("Tak", type.Name);
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

