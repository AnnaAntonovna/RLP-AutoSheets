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

namespace RLP
{
    [Transaction(TransactionMode.Manual)]
    public class Combine : IExternalCommand
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

            //var window = new BlockingWindow();
            //window.Show("Processing...");

            try
            {
                AssemblyInstance assemblyInstance = doc.GetElement(assemblyRef) as AssemblyInstance;

                if (assemblyInstance != null)
                {
                    //INFOOO

                    string assebmlyName = assemblyInstance.Name;

                    List<ViewSheet> sheets = new List<ViewSheet>(); // list of sheets
                    List<ViewSection> facades = new List<ViewSection>(); ; // list of lists of facades, where each inner list corresponds to a sheet in the sheets list


                    /////SHEEEEEEEEEEEEEETTSS
                    ///
                    using (Transaction tx = new Transaction(doc, "CreateSheets"))
                    {
                        tx.Start();
                        ViewSheet sheet1;
                        try
                        {
                            if (assebmlyName.Contains("3НСН"))
                            {
                                sheet1 = CreateNewSheet(doc, assebmlyName, "Форма 4_Панели (масштаб 1 к 20)", 2, $"Наружная стеновая панель {assebmlyName}", "1"); sheets.Add(sheet1);
                                ViewSheet sheet2 = CreateNewSheet(doc, assebmlyName, "Форма 6", 2, $"Наружная стеновая панель {assebmlyName}. Армирование внутреннего слоя", "2"); sheets.Add(sheet2);
                                ViewSheet sheet3 = CreateNewSheet(doc, assebmlyName, "Форма 6", 3, $"Наружная стеновая панель {assebmlyName}. Армирование внешнего слоя", "3"); sheets.Add(sheet3);
                                ViewSheet sheet4 = CreateNewSheet(doc, assebmlyName, "Форма 6", 3, $"Наружная стеновая панель {assebmlyName}. Схема рустовки", "4"); sheets.Add(sheet4);
                                ViewSheet sheet5 = CreateNewSheet(doc, assebmlyName, "Форма 6", 2, $"Наружная стеновая панель {assebmlyName}. Спецификации", "5"); sheets.Add(sheet5);
                                tx.Commit();
                                uiDoc.ActiveView = sheet1;
                            }
                            else if (assebmlyName.Contains("1НСН"))
                            {
                                ViewSheet sheet6 = CreateNewSheet(doc, assebmlyName, "Форма 4_Панели (масштаб 1 к 20)", 2, $"Наружная стеновая панель {assebmlyName}", "1"); sheets.Add(sheet6);
                                ViewSheet sheet7 = CreateNewSheet(doc, assebmlyName, "Форма 6", 2, $"Наружная стеновая панель {assebmlyName}. Армирование", "2"); sheets.Add(sheet7);
                                ViewSheet sheet8 = CreateNewSheet(doc, assebmlyName, "Форма 6", 2, $"Наружная стеновая панель {assebmlyName}. Спецификации", "3"); sheets.Add(sheet8);
                                tx.Commit();
                            }
                            else if (assebmlyName.Contains("ПСВ"))
                            {
                                ViewSheet sheet9 = CreateNewSheet(doc, assebmlyName, "Форма 4_Панели (масштаб 1 к 20)", 2, $"Внутренняя стеновая панель {assebmlyName}", "1"); sheets.Add(sheet9);
                                ViewSheet sheet10 = CreateNewSheet(doc, assebmlyName, "Форма 6", 2, $"Внутренняя стеновая панель {assebmlyName}. Армирование", "2"); sheets.Add(sheet10);
                                ViewSheet sheet11 = CreateNewSheet(doc, assebmlyName, "Форма 6", 2, $"Внутренняя стеновая панель {assebmlyName}. Спецификации", "3"); sheets.Add(sheet11);
                                tx.Commit();
                            }
                            else
                            { TaskDialog.Show("Oh..", "Название сборки не корректное. Листы не созданы"); tx.RollBack(); }
                        }
                        catch (Exception ex) { ShowException("Can't create all sheets", ex); tx.RollBack(); }
                    }

                    ////VIEEESSSSSSSSSSSSSS

                    Wall hostWall = null;
                    ViewSection sectionView = null;

                    //METHODS

                    try { hostWall = GetWallHostForFirstRebarInAssembly(assemblyInstance); }
                    catch (Exception ex) { ShowException("Can't get a wall", ex); return Result.Failed; }

                    //FACADES!!!!


                    //All Facades

                    string viewportTypeName = "Заголовок на листе";

                    //Templates

                    string viewTemplateName1 = "3НСНг_Панель_Фасад_Опалубка";
                    string viewTemplateName2 = "3НСНг_Панель_Фасад_Армирование внутреннего слоя";
                    string viewTemplateName3 = "3НСНг_Панель_Фасад_Армирование внешнего слоя";
                    string viewTemplateName4 = "3НСНг_Панель_Фасад_Рустовка";

                    string viewTemplateName56 = "НСН-ПСВ_Панель_Фасад_Опалубка";
                    string viewTemplateName78 = "НСН-ПСВ_Панель_Фасад_Армирование внутреннего слоя";

                    using (Transaction tx = new Transaction(doc, "CreateFacades"))
                    {
                        //tx.Start();

                        if (assebmlyName.Contains("3НСН"))
                        {
                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Опалубка", viewTemplateName1, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade1", ex); return Result.Failed; }

                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Армирование внутреннего слоя", viewTemplateName2, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade2", ex); return Result.Failed; }

                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Армирование внешнего слоя", viewTemplateName3, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade3", ex); return Result.Failed; }

                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Схема рустовки", viewTemplateName4, viewportTypeName, false); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade4", ex); return Result.Failed; }

                        }
                        else if (assebmlyName.Contains("1НСН"))
                        {
                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Опалубка", viewTemplateName56, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade1", ex); return Result.Failed; }

                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Наружная стеновая панель {assebmlyName}. Армирование", viewTemplateName78, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade2", ex); return Result.Failed; }
                        }
                        else if (assebmlyName.Contains("ПСВ"))
                        {
                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Внутренняя стеновая панель {assebmlyName}. Опалубка", viewTemplateName56, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade1", ex); return Result.Failed; }

                            try
                            { sectionView = (ViewSection)GetWonderfulFacade(hostWall, assemblyInstance, $"Внутренняя стеновая панель {assebmlyName}. Армирование", viewTemplateName78, viewportTypeName, true); facades.Add(sectionView); }
                            catch (Exception ex) { ShowException("Can't GetWonderfulFacade2", ex); return Result.Failed; }
                        }
                        
                       // tx.Commit();
                    }

                    ////SECTIONS 
                    ///
                    string viewportTypeNameSection = "Сечение_Номер вида";

                    //Templates

                    string viewTemplateNameOpalubka = "3НСНг_Панель_Сечение_Опалубка";
                    string viewTemplateNameArmirovanie = "3НСНг_Панель_Сечение_Армирование";
                    string viewTemplateNameRustovka = "3НСНг_Панель_Разрез_Рустовка";


                    string viewTemplateNameVnOpalubka = "НСН-ПСВ_Панель_Сечение_Опалубка";
                    string viewTemplateNameVnArm = "НСН-ПСВ_Панель_Сечение_Армирование";

                    using (Transaction tx = new Transaction(doc, "Create section views"))
                    {
                        tx.Start();
                        try
                        {
                            if (assebmlyName.Contains("3НСН"))
                            {
                                int sectionNumber = 2; //number amomg all viewports on a sheet

                                //OPALUBKA
                                List<XYZ> horPoints = GetPointsForHorizontalSections(hostWall);
                                int horSectionNumber1 = 1;
                                foreach (XYZ horPoint in horPoints)
                                {
                                    if (horPoint == horPoints[0])
                                    {

                                    }
                                    else
                                    {
                                        ViewSection opalHorView = CreateVerticalSectionView(hostWall, horPoint);
                                        SetParameters(opalHorView, $"{sectionNumber}", $"О_Разрез_Опл горизонтальный {horSectionNumber1}", viewTemplateNameOpalubka, viewportTypeNameSection);
                                        sectionNumber++;
                                        horSectionNumber1++;
                                    }
                                }

                                List<XYZ> vertPoints = GetPointsForVerticalSections(hostWall, false);
                                int vertSectionNumberOp = 1; // number among only vert viewportss on the sheet
                                foreach (XYZ point in vertPoints)
                                {

                                    ViewSection opalVertView = CreateVerticalSectionView(hostWall, point);

                                    SetParameters(opalVertView, $"{sectionNumber}", $"О_Разрез_Опл вертикальный {vertSectionNumberOp}", viewTemplateNameOpalubka, viewportTypeNameSection);
                                    vertSectionNumberOp++;
                                }

                                //ARM

                                int vertSectionNumberAr = 1; // number among only vert viewportss on the sheet
                                foreach (XYZ point in vertPoints)
                                { 
                                    ViewSection armVertView = CreateVerticalSectionView(hostWall, point);
                                    SetParameters(armVertView, $"{sectionNumber}", $"О_Разрез_Арм вертикальный {vertSectionNumberAr}", viewTemplateNameOpalubka, viewportTypeNameSection);
                                    vertSectionNumberAr++;

                                }
                            }
                            else if (assebmlyName.Contains("1НСН"))
                            {
                                List<XYZ> vertPointsOp = GetPointsForVerticalSections(hostWall, true);

                                List<XYZ> vertPointsArm = GetPointsForVerticalSections(hostWall, false);

                            }
                            else if (assebmlyName.Contains("ПСВ")) 
                            {
                                List<XYZ> vertPointsOp = GetPointsForVerticalSections(hostWall, true);

                                List<XYZ> vertPointsArm = GetPointsForVerticalSections(hostWall, false);

                            }
                        }
                        catch { }
                    }



                            ///PLACES 
                            ///

                    using (Transaction tx = new Transaction(doc, "Place views on sheets"))
                    {
                        tx.Start();
                        try
                        {
                            foreach (ViewSheet sheet in sheets)
                            {
                                int index = sheets.IndexOf(sheet);
                                if (index < facades.Count())
                                {
                                    ViewSection facade = facades[index];
                                    // check if facade is already on the sheet

                                    if (!sheet.GetAllPlacedViews().Contains(facade.Id))
                                    {
                                        // Create a new viewport for the facade
                                        BoundingBoxXYZ boundingBox = sheet.get_BoundingBox(null);
                                        XYZ minPoint = boundingBox.Min;
                                        XYZ maxPoint = boundingBox.Max;
                                        Viewport viewport = Viewport.Create(sheet.Document, sheet.Id, facade.Id, XYZ.Zero);

                                        //TaskDialog.Show("Size", (sheet.Outline.Max - sheet.Outline.Min).GetLength().ToString());

                                        SetViewportType(facade, "Заголовок на листе");

                                        XYZ vector;

                                        if ((sheet.Outline.Max - sheet.Outline.Min).GetLength() > 2)
                                        {
                                            vector = new XYZ((-sheet.Outline.Max.U + sheet.Outline.Min.U) * 0.7, (+sheet.Outline.Max.V - sheet.Outline.Min.V) * 0.75, 0);
                                        }
                                        else 
                                        {
                                            vector = new XYZ((-sheet.Outline.Max.U + sheet.Outline.Min.U) * 0.6, (+sheet.Outline.Max.V - sheet.Outline.Min.V) * 0.63, 0);
                                        }
                                        
                                        viewport.Location.Move(vector);
                                        BoundingBoxXYZ viewportBox = viewport.get_BoundingBox(sheet);

                                        // calculate the center points
                                        XYZ vectorLabel = new XYZ((facade.Outline.Max - facade.Outline.Min).U*0.5, (facade.Outline.Max - facade.Outline.Min).V*0.98, 0);

                                        //viewport.LabelOffset.Add(vectorLabel);
                                        viewport.LabelOffset = vectorLabel;
                                    }
                                }
                            }
                        }
                        catch (Exception ex) { ShowException("Can't place views", ex); }
                        tx.Commit();
                    }
                    foreach (ViewSheet sheet in sheets)
                    {
                        uiDoc.ActiveView = sheet;
                    }


                    return Result.Succeeded;
                }
                //window.Close();
                TaskDialog.Show("Некорректная сборка", "Операция отменена"); 
                return Result.Cancelled;
                
            }
            catch (Exception ex) { ShowException("BIG MISTAKE / Assembly", ex); return Result.Failed; }
        }
        public ViewSection GetWonderfulFacade(Wall hostWall, AssemblyInstance assembly, string viewName, string viewTemplateName, string viewportTypeName, bool lookInside)
        {
            string assemblyName = assembly.Name;
            try
            {
                if (new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Any(x => x.Name == viewName))
                {
                    return new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(x => x.Name == viewName).First();
                }

                ViewSection sectionView = null;
                // Create a wall section view
                using (Transaction tx = new Transaction(doc, "CreateWallSectionView"))
                {
                    tx.Start();
                    try { sectionView = CreateWallSectionView(hostWall, lookInside); }
                    catch (Exception ex) { ShowException("Can't CreateWallSectionView", ex);  return null; }
                    tx.Commit();
                }

                using (Transaction tx = new Transaction(doc, "SetParameters"))
                {
                    tx.Start();
                    // Set the parameters of the view
                    try { SetParameters(sectionView, assemblyName, viewName, viewTemplateName, viewportTypeName); }
                    catch (Exception ex) { ShowException("Can't SetParameters", ex); return null; }
                    tx.Commit();
                }

                using (Transaction tx = new Transaction(doc, "HideElements"))
                {
                    tx.Start();
                    try { HideElementsInView(sectionView, hostWall, assembly); }
                    catch (Exception ex) { ShowException("Can't HideElementsInView", ex); return null; }
                    tx.Commit();
                }

                using (Transaction tx = new Transaction(doc, "TagElementsInViewOnCenter"))
                {
                    tx.Start();
                    // Tag the elements in the view on their center
                    //try { TagElementsInViewOnCenter(sectionView, assembly); }
                    //catch (Exception ex) { ShowException("Can't  TagElementsInViewOnCenter", ex); return null; }
                    tx.Commit();
                }
                return sectionView;
            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't GetWonderfulFacade");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
                return null;
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
        public void ShowException(string errorName, Exception ex)
        {
            TaskDialog taskDialog = new TaskDialog("Oh-oh, " + errorName);
            taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
            taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
            taskDialog.Show();
        }
        public static Wall GetWallHostForFirstRebarInAssembly(AssemblyInstance assembly)
        {
            // Get all the subelements of the assembly
            IList<ElementId> subElementIds = (IList<ElementId>)assembly.GetMemberIds();
            try
            {
                // Loop through the subelements until we find a rebar
                foreach (ElementId subElementId in subElementIds)
                {
                    Element subElement = assembly.Document.GetElement(subElementId);
                    if (subElement is Rebar)
                    {
                        // We found a rebar! Get its host element
                        ElementId hostElementId = (subElement as Rebar).GetHostId();
                        Element hostElement = assembly.Document.GetElement(hostElementId);

                        // Make sure the host element is a wall
                        if (hostElement is Wall)
                        {
                            return hostElement as Wall;
                        }
                    }
                }
            }
            catch (Exception ex) { TaskDialog.Show("Wall method exception", ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite); }
            // If we got here, we didn't find a rebar or its host element
            return null;
        }
        private ViewSection CreateWallSectionView(Wall wall, bool lookAtNakedWall)
        {
            Document doc = wall.Document;
            ViewFamilyType vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault<ViewFamilyType>(x =>
                    ViewFamily.Section == x.ViewFamily);
            double lookRatio = 1;
            if (lookAtNakedWall) { lookRatio = -1; }

            // Determine section box
            LocationCurve lc = wall.Location as LocationCurve;
            Line line = lc.Curve as Line;
            if (null == line)
            {
                throw new Exception("Unable to retrieve wall location line.");
            }

            XYZ up = XYZ.BasisZ;
            //XYZ viewdir = XYZ wallNormal = line.Direction.CrossProduct(XYZ.BasisZ).Normalize();
            XYZ walldir = line.Direction.Normalize();
            XYZ viewdir = walldir.CrossProduct(up);

            walldir = viewdir.CrossProduct(up);
            viewdir = lookRatio * wall.Orientation; //!!!!!!!!!!!!!!!!! And to multiply ot by parameter (-1) изнути 1 снаружи!!!
            walldir = viewdir.CrossProduct(up);
            viewdir = walldir.CrossProduct(up);


            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ v = q - p;

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            double minZ = bb.Min.Z;
            double maxZ = bb.Max.Z;

            double w = v.GetLength();
            double h = maxZ - minZ;
            double d = wall.WallType.Width;
            double offset = 0.1 * w;

            XYZ min = new XYZ(-w * 0.5 - offset, minZ - offset, -offset);
            XYZ max = new XYZ(w * 0.5 + offset, maxZ + offset, 0);

            XYZ midpoint = p + 0.5 * v;
            //walldir = v.Normalize();
            up = XYZ.BasisZ;
            viewdir = walldir.CrossProduct(up);
            BoundingBoxXYZ sectionBox = bb;

            Transform t = Transform.Identity;
            t.Origin = new XYZ(midpoint.X, midpoint.Y, 0);//midpoint;
            t.BasisX = walldir;
            t.BasisY = up;
            t.BasisZ = viewdir;

            if (sectionBox == null)
            {
                throw new Exception("Unable to retrieve wall bounding box.");
            }

            //TaskDialog.Show("Debug", "sectionBox before transform: Min = " + sectionBox.Min + ", Max = " + sectionBox.Max + ", Transform = " + sectionBox.Transform);
            //TaskDialog.Show("Transform Information", t.Origin.ToString());


            try { sectionBox.Transform = t; }
            catch (Exception ex) { ShowException("Can't transorm", ex); }
            sectionBox.Min = min;
            sectionBox.Max = max;

            // Create wall section view

            ViewSection sectionView = ViewSection.CreateSection(doc, vft.Id, sectionBox);
            return sectionView;

        }
       
        public void SetParameters(ViewSection view, string assemblyName, string viewName, string viewTemplateName, string viewportTypeName)
        {
            try
            {
                // Set the view name and viewport detail number parameters
                view.get_Parameter(BuiltInParameter.VIEW_NAME).Set(viewName);
                view.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(assemblyName);

                // Apply the view template
                ApplyViewTemplate(view, viewTemplateName);

                // Set the viewport type
                
                SetViewportType(view, viewportTypeName);
            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't SetParameters");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
            }
        }

        private void ApplyViewTemplate(ViewSection view, string viewTemplateName)
        {
            // Get the view template element
            View viewTemplate = new FilteredElementCollector(view.Document)
                .OfClass(typeof(ViewSection))
                .Cast<ViewSection>()
                .FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

            if (viewTemplate == null)
            {
                TaskDialog.Show("Error", $"The view template named {viewTemplateName} does not exist.");
                return;
            }
            try
            {
                // Set the view template parameter of the view
                
                 view.ViewTemplateId = viewTemplate.Id;

            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't ApplyViewTemplate");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
            }
        }


        public void TagElementsInViewOnCenter(ViewSection view, AssemblyInstance assembly)
        {
            // Get all the elements visible in the view
            IList<ElementId> collector = (IList<ElementId>)new FilteredElementCollector(view.Document, view.Id).ToElementIds();
            IList<ElementId> subElementIds = (IList<ElementId>)assembly.GetMemberIds();

            IList<ElementId> commonElements = collector.Intersect(subElementIds).ToList();


            // Loop through the elements and tag each one on its center
            foreach (ElementId elementId in commonElements)
            {
                try
                {
                    Element element = doc.GetElement(elementId);

                    //if (element.Category != null && element.Category.Name == "Rebar")
                    //    continue;

                    // Get the center point of the element
                    BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                    XYZ center = (bbox.Min + bbox.Max) / 2.0;

                    // Create a new tag
                    IndependentTag tag = IndependentTag.Create(view.Document, view.Id, new Reference(element), false, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, center);

                    // Set the tag leader to point to the center of the element
                    XYZ leaderEnd = center + new XYZ(1, 1, 0);

                    tag.LeaderEndCondition = LeaderEndCondition.Attached;
                    //tag.LeaderEnd = leaderEnd;
                }
                catch (Exception ex)
                {
                    
                }
            }
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
                            viewport.ChangeTypeId(scaledViewPortType.Id);
                            doc.Regenerate();
                        }
                        catch(Exception ex)
                        {
                            ShowException("Can't change ViewPort", ex); 
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

        public static List<Viewport> GetViewports(View view)
        {
            var collector = new FilteredElementCollector(view.Document)
                .OfClass(typeof(Viewport))
                .WhereElementIsNotElementType()
                .Cast<Viewport>();

            var viewports = collector.Where(vp => vp.ViewId == view.Id).ToList();
            return viewports;
        }


        public ViewSheet CreateNewSheet(Document doc, string assemblyName, string titleBlockName,
            int aForm, string sheetName, string sheetNumber)
        {
            try
            {

                //check if a sheet exists

                if (new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().Any(x => x.Name == sheetName))
                {
                    return new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().Where(x => x.Name == sheetName).First();
                }

                // Get the title block type id for the desired type
                FilteredElementCollector titleBlocks = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_TitleBlocks);
                FamilySymbol titleBlock = titleBlocks
                    .Where(x => x.Name == titleBlockName).FirstOrDefault() as FamilySymbol;
                ElementId titleBlockTypeId = titleBlock.Id;

                // Create the new sheet
                ViewSheet newSheet = ViewSheet.Create(doc, titleBlockTypeId);

                newSheet.Name = sheetName;

                // TITLE BLOCK ON THE SHEET (ШТАМП)

                Element titleBlockInstance = new FilteredElementCollector(doc, newSheet.Id)
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
                Parameter nomerLista = newSheet.LookupParameter("RLP_Номер листа");
                if (nomerLista != null)
                {
                    nomerLista.Set(sheetNumber);
                }
                else { TaskDialog.Show("Can't set", $"RLP_Номер листа does not set"); }



                //Set common parameters (1 method), use three types of it

                if (assemblyName.Contains("3НСН"))
                {
                    SetSheetParameterValues(newSheet, $"КЖ1.И1-{assemblyName}",
                    "Конструкции железобетонных и арматурных изделий. Надземная часть. Панели наружные",
                     5);
                }
                else if (assemblyName.Contains("1НСН"))
                {
                    SetSheetParameterValues(newSheet, $"КЖ1.И1-{assemblyName}",
                            "Конструкции железобетонных и арматурных изделий. Надземная часть. Панели наружные",
                             3);
                }
                else if (assemblyName.Contains("ПСВ"))
                {
                    SetSheetParameterValues(newSheet, $"КЖ1.И2-{assemblyName}",
                            "Конструкции железобетонных и арматурных изделий. Надземная часть. Панели внутренни",
                             3);
                }
                else
                { TaskDialog.Show("Oh..", "Название сборки не корректное. Параметры листа не заполнены. Пожалуйста, проверьте название сборки."); }

                return newSheet;
            }
            catch (Exception ex) { ShowException("Can't CreateNewSheet", ex); return null; }
        }


        private void SetSheetParameterValues(ViewSheet sheet, string razdel, string naimenovanie, int kolvo)
        {
            try
            {
                // Get the document
                Document doc = sheet.Document;

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


        public void HideElementsInView(View view, Element hostWall, AssemblyInstance assembly)
        {
            Document doc = view.Document;


            // Collect all elements in the view except the hostWall and assembly
            FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id);
            ICollection<ElementId> elementIdsToHide = collector
                .Where(e => e.Name != assembly.Name && e.Id != view.Id && e.CanBeHidden(view)
                && !assembly.GetMemberIds().ToList().Contains(e.Id))
                .Select(e => e.Id)
                .ToList();

            // Hide all elements except the hostWall and assembly
            view.HideElements(elementIdsToHide);
        }

        public List<XYZ> GetPointsForVerticalSections(Wall wall, bool opalVnutr)
        {

            LocationCurve lc = wall.Location as LocationCurve;
            Autodesk.Revit.DB.Line line = lc.Curve as Line;
            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ v = q - p;

            List<XYZ> centerPoints = new List<XYZ>();

            // Get all the openings in the wall
            IList<ElementId> openingIds = wall.FindInserts(true, false, false, false);

            if (opalVnutr)
            {
                XYZ midpoint = p + 0.1 * v;
                centerPoints.Add(midpoint);
                return centerPoints;
            }

            if (openingIds == null || openingIds.Count == 0)
            {
                XYZ midpoint = p + 0.5 * v;
                centerPoints.Add(midpoint);
                return centerPoints;
            }
            // Get the unique family and type names of the openings  HERE THE MISTAKE!
            IList<FamilyInstance> openings = new FilteredElementCollector(doc, openingIds)
               .OfClass(typeof(FamilyInstance))
               .Cast<FamilyInstance>()
               .Where(o => o.Symbol.Category != null && !o.Symbol.Category.Name.Contains("Generic Models"))
               .ToList();

            if (openings.Count() > 0)
            {
                XYZ midpoint = p + 0.1 * v;
                centerPoints.Add(midpoint);

            }
            else
            {
                XYZ midpoint = p + 0.5 * v;
                centerPoints.Add(midpoint);
                return centerPoints;
            }

            var uniqueFamilies = openings
                .Select(fi => fi.Symbol.Family.Name)
                .Distinct();

            var uniqueFamilyTypes = openings
                .Select(fi => fi.Symbol.Name)
                .Distinct();

            // Get the center points of the openings with unique family and type names
            foreach (string familyName in uniqueFamilies)
            {
                foreach (string typeName in uniqueFamilyTypes)
                {
                    Element opening = openings
                        .Where(fi => fi.Symbol.Family.Name == familyName
                        && fi.Symbol.Name == typeName)
                        .FirstOrDefault() as Element;
                    if (opening != null)
                    {
                        XYZ centerPoint = GetBoundingBoxCenter(opening);
                        if (familyName.Contains("+"))
                        {
                            // Change the offset value as per your requirements

                            double offsetDistance = 0.15 * line.Length;
                            XYZ direction = line.Direction.Normalize();
                            XYZ offsetXYZ = direction.Multiply(offsetDistance);

                            XYZ walloffset = 0.1 * (line.Length) * direction;

                            XYZ centerPointPlus = centerPoint + offsetXYZ;
                            XYZ centerPointMinus = centerPoint - offsetXYZ;

                            ///To put points to the footing of a wall

                            XYZ cPplus = new XYZ(centerPointPlus.X, centerPointPlus.Y, p.Z);
                            XYZ cPminus = new XYZ(centerPointMinus.X, centerPointMinus.Y, p.Z);


                            centerPoints.Add(cPplus);
                            centerPoints.Add(cPminus);


                            if ((centerPoints[0] - cPminus).GetLength() < offsetDistance)
                            {
                                centerPoints[0] = q - walloffset;
                            }
                            if ((cPplus - centerPoints[0]).GetLength() < offsetDistance)
                            {
                                centerPoints[0] = p + walloffset;
                            }
                        }
                        else
                        {
                            centerPoints.Add(new XYZ(centerPoint.X, centerPoint.Y, p.Z));
                        }
                    }
                }
            }

            return centerPoints;
        }



        public List<XYZ> GetPointsForHorizontalSections(Wall wall, double offset = 0.0)
        {

            LocationCurve lc = wall.Location as LocationCurve;
            Autodesk.Revit.DB.Line line = lc.Curve as Line;
            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ v = q - p;

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            double minZ = bb.Min.Z;
            double maxZ = bb.Max.Z;

            double h = maxZ - minZ;

            XYZ midpoint = p + 0.5 * v;
            XYZ cpointCenter = new XYZ(midpoint.X, midpoint.Y, midpoint.Z + h * 0.5);
            XYZ cpointDoor = new XYZ(midpoint.X, midpoint.Y, midpoint.Z + h * 0.2);

            List<XYZ> centerPoints = new List<XYZ>();
            // Get all the openings in the wall
            IList<ElementId> openingIds = wall.FindInserts(true, false, false, false);

            XYZ pointLookUp = new XYZ(midpoint.X, midpoint.Y, midpoint.Z + h * 0.1);
            centerPoints.Add(pointLookUp);

            if (openingIds == null || openingIds.Count == 0)
            {
                centerPoints.Add(cpointCenter);
                return centerPoints;
            }

            // Get the unique family and type names of the openings
            IList<FamilyInstance> openings = new FilteredElementCollector(doc, openingIds)
               .OfClass(typeof(FamilyInstance))
               .Cast<FamilyInstance>()
               .Where(o => o.Symbol.Category != null && !o.Symbol.Category.Name.Contains("Generic Models"))
               .ToList();

            if (openings.Count() < 1)
            {
                centerPoints.Add(cpointCenter);
                return centerPoints;
            }

            var uniqueFamilies = openings
                .Select(fi => fi.Symbol.Family.Name)
                .Distinct();

            var uniqueFamilyTypes = openings
                .Select(fi => fi.Symbol.Name)
                .Distinct();


            // Get the center points of the openings with unique family and type names

            foreach (string familyName in uniqueFamilies)
            {
                foreach (string typeName in uniqueFamilyTypes)
                {
                    Element opening = openings
                        .Where(fi => fi.Symbol.Family.Name == familyName
                        && fi.Symbol.Name == typeName)
                        .FirstOrDefault() as Element;
                    if (opening != null)
                    {
                        XYZ centerPoint = GetBoundingBoxCenter(opening);
                        if (familyName.Contains("+"))
                        {

                            if (!centerPoints.Contains(cpointCenter))
                            { centerPoints.Add(cpointCenter); }
                            if (!centerPoints.Contains(cpointDoor))
                            { centerPoints.Add(cpointDoor); }
                        }
                        else
                        {
                            if (openings.Count() > 1)
                            {
                                if (!centerPoints.Contains(cpointCenter))
                                { centerPoints.Add(cpointCenter); }
                                else
                                {
                                    if (!centerPoints.Contains(cpointDoor))
                                    { centerPoints.Add(cpointDoor); }
                                }
                            }
                            if (openings.Count() == 1)
                            {
                                if (!centerPoints.Contains(cpointCenter))
                                { centerPoints.Add(cpointCenter); }
                            }
                        }

                    }
                }
            }

            return centerPoints;
        }




        public XYZ GetBoundingBoxCenter(Element elem)
        {
            BoundingBoxXYZ boundingBox = elem.get_BoundingBox(null);
            XYZ min = boundingBox.Min;
            XYZ max = boundingBox.Max;
            XYZ center = (min + max) / 2.0;
            return center;
        }




        private ViewSection CreateVerticalSectionView(Wall wall, XYZ originPoint)
        {
            Document doc = wall.Document;
            ViewFamilyType vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault<ViewFamilyType>(x =>
                    ViewFamily.Section == x.ViewFamily);

            // Determine section box
            LocationCurve lc = wall.Location as LocationCurve;
            Line line = lc.Curve as Line;
            if (null == line)
            {
                throw new Exception("Unable to retrieve wall location line.");
            }

            XYZ yView = XYZ.BasisZ;

            XYZ xView = wall.Orientation; //!!!!!!!!!!!!!!!!! And to multiply ot by parameter (-1) изнути 1 снаружи!!!
            XYZ viewdir = xView.CrossProduct(yView);
            xView = viewdir.CrossProduct(yView);
            viewdir = xView.CrossProduct(yView);

            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ v = q - p;

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);

            //walldir = v.Normalize();
            BoundingBoxXYZ sectionBox = bb;

            Transform t = Transform.Identity;
            t.Origin = new XYZ(originPoint.X, originPoint.Y, 0);//midpoint;
            t.BasisX = xView;
            t.BasisY = yView;
            t.BasisZ = viewdir;

            if (sectionBox == null)
            {
                throw new Exception("Unable to retrieve wall bounding box.");
            }

            try { sectionBox.Transform = t; }
            catch (Exception ex) { ShowException("Can't transorm", ex); }

            double minZ = bb.Min.Z;
            double maxZ = bb.Max.Z;

            double w = v.GetLength();
            double h = maxZ - minZ;
            double d = wall.WallType.Width;
            double offset = 0.1 * h;

            XYZ min = new XYZ(-0.5 * d - offset, minZ - offset, -d);
            XYZ max = new XYZ(d * 0.5 + offset, maxZ + offset, 0);

            sectionBox.Min = min;
            sectionBox.Max = max;

            // Create wall section view

            ViewSection sectionView = ViewSection.CreateSection(doc, vft.Id, sectionBox);
            return sectionView;

        }




        private ViewSection CreateHorizontalSectionView(Wall wall, XYZ originPoint, bool upDown)
        {

            double look = 1;
            if (upDown)
            { look = 1; }
            else
            { look = -1; }
            Document doc = wall.Document;
            ViewFamilyType vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault<ViewFamilyType>(x =>
                    ViewFamily.Section == x.ViewFamily);

            // Determine section box
            LocationCurve lc = wall.Location as LocationCurve;
            Line line = lc.Curve as Line;
            if (null == line)
            {
                throw new Exception("Unable to retrieve wall location line.");
            }

            XYZ viewdir = XYZ.BasisZ * look;

            XYZ yView = wall.Orientation;

            XYZ xView = yView.CrossProduct(viewdir);
            yView = xView.CrossProduct(viewdir);
            xView = yView.CrossProduct(viewdir);

            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ v = q - p;

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);

            //walldir = v.Normalize();
            BoundingBoxXYZ sectionBox = bb;

            Transform t = Transform.Identity;
            t.Origin = new XYZ(originPoint.X, originPoint.Y, originPoint.Z);//midpoint;
            t.BasisX = xView;
            t.BasisY = yView;
            t.BasisZ = viewdir;

            if (sectionBox == null)
            {
                throw new Exception("Unable to retrieve wall bounding box.");
            }

            try { sectionBox.Transform = t; }
            catch (Exception ex) { ShowException("Can't transorm", ex); }

            double minZ = bb.Min.Z;
            double maxZ = bb.Max.Z;

            double w = v.GetLength();
            double h = maxZ - minZ;
            double d = wall.WallType.Width;
            double offset = 0.1 * h;

            XYZ min = new XYZ(-0.5 * w - offset, -0.5 * d - offset, -d);
            XYZ max = new XYZ(w * 0.5 + offset, d * 0.5 + offset, 0);

            sectionBox.Min = min;
            sectionBox.Max = max;

            // Create wall section view

            ViewSection sectionView = ViewSection.CreateSection(doc, vft.Id, sectionBox);
            return sectionView;

        }
    }
}
