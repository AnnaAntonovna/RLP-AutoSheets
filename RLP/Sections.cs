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
using System.Windows.Controls;

namespace RLP
{
    [Transaction(TransactionMode.Manual)]
    public static class SectionMethods
    {
        public static List<XYZ> allVerticalPoints(Wall wall, Document doc)
        {
            List<XYZ> verticalPoints = GetPointsForVerticalSections(wall, false, doc);
            return verticalPoints;
        }
        public static XYZ leftVerticalSectionPoint(Wall wall, Document doc)
        {
            List<XYZ> uniqueVerticalPoints = GetPointsForVerticalSections(wall, true, doc);
            return uniqueVerticalPoints.First();
        }

        public static List<XYZ> horizontalPoints12(Wall wall, Document doc)
        {
            List<XYZ> horizontalPoints = GetPointsForHorizontalSections(wall, doc).Skip(0).ToList();
            return horizontalPoints;
        }

        public static XYZ horizontalPoint0(Wall wall, Document doc)
        {
            XYZ horizontalPoint = GetPointsForHorizontalSections(wall, doc).First();
            return horizontalPoint;
        }

        public static XYZ horizontalPoint1center(Wall wall, Document doc)
        {
            XYZ horizontalPoint = GetPointsForHorizontalSections(wall, doc)[1];
            return horizontalPoint;
        }

        public static ViewSection GetHorizontalSection(bool lookDown, string viewPortNumber, XYZ originPoint, Wall hostWall, AssemblyInstance assembly, string viewName, string viewTemplateName, string viewportTypeName,Document doc)
        {
            try
            {
                if (new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Any(x => x.Name == viewName))
                {
                    return new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(x => x.Name == viewName).First();
                }

                ViewSection sectionView = null;
                // Create a wall section view
                using (Transaction tx = new Transaction(doc, "CreateHorizontalSectionView"))
                {
                    tx.Start();
                    try { sectionView = CreateHorizontalSectionView(hostWall, originPoint, lookDown, viewName); }
                    catch (Exception ex) { ShowException("Can't CreateHorizontalSectionView", ex); return null; }
                    tx.Commit();
                }

                using (Transaction tx = new Transaction(doc, "SetParameters"))
                {
                    tx.Start();
                    // Set the parameters of the view
                    try { SetSectionParameters(sectionView, viewPortNumber, viewName, viewTemplateName, viewportTypeName, doc); }
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


        public static ViewSection GetVerticalSection(XYZ originPoint, string viewPortNumber, Wall hostWall, AssemblyInstance assembly, string viewName, string viewTemplateName, string viewportTypeName, Document doc)
        {
            try
            {
                if (new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Any(x => x.Name == viewName))
                {
                    return new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(x => x.Name == viewName).First();
                }

                ViewSection sectionView = null;
                // Create a wall section view
                using (Transaction tx = new Transaction(doc, "CreateVerticalSectionView"))
                {
                    tx.Start();
                    try { sectionView = CreateVerticalSectionView(hostWall, originPoint, viewName); }
                    catch (Exception ex) { ShowException("Can't CreateVerticalSectionView", ex); return null; }
                    tx.Commit();
                }

                using (Transaction tx = new Transaction(doc, "SetParameters"))
                {
                    tx.Start();
                    // Set the parameters of the view
                    try { SetSectionParameters(sectionView, viewPortNumber, viewName, viewTemplateName, viewportTypeName, doc); }
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


        public static List<XYZ> GetPointsForVerticalSections(Wall wall, bool unique, Document doc)
        {

            LocationCurve lc = wall.Location as LocationCurve;
            Autodesk.Revit.DB.Line line = lc.Curve as Line;
            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ v = q - p;

            List<XYZ> centerPoints = new List<XYZ>();

            // Get all the openings in the wall
            IList<ElementId> openingIds = wall.FindInserts(true, false, false, false);

            if (unique)
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
                XYZ midpoint = p + 0.05 * v;
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

        public static bool IsViewOnSheet(View targetView, ViewSheet targetSheet)
        {
            // Get all the viewports on the target sheet
            IEnumerable<Viewport> viewports = new FilteredElementCollector(targetSheet.Document)
                .OfCategory(BuiltInCategory.OST_Viewports)
                .Cast<Viewport>();

            // Check if the target view is already placed on the sheet
            bool isViewOnSheet = viewports.Any(vp => vp.ViewId == targetView.Id);

            if (isViewOnSheet)
            {
                TaskDialog.Show("Вид уже есть на листе", $"Вид {targetView.Name} уже расположен на листе {targetSheet.Name}");
            }

            return isViewOnSheet;
        }

        public static List<XYZ> GetPointsForHorizontalSections(Wall wall, Document doc)
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
            XYZ cpointCenter = new XYZ(midpoint.X, midpoint.Y, midpoint.Z + h * 0.4);
            XYZ cpointDoor = new XYZ(midpoint.X, midpoint.Y, midpoint.Z + h * 0.15);

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

        private static void ApplyViewTemplate(ViewSection view, string viewTemplateName)
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
        public static void HideElementsInView(View view, Element hostWall, AssemblyInstance assembly)
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

        public static void SetSectionParameters(ViewSection view, string viewportNumber, string viewName, string viewTemplateName, string viewportTypeName, Document doc)
        {
            try
            {
                // Set the view name and viewport detail number parameters
                view.get_Parameter(BuiltInParameter.VIEW_NAME).Set(viewName);
                view.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(viewportNumber);

                // Apply the view template
                ApplyViewTemplate(view, viewTemplateName);

                // Set the viewport type

                SetViewportType(view, viewportTypeName, doc);

            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't SetParameters");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
            }
        }

        private static ViewSection CreateVerticalSectionView(Wall wall, XYZ originPoint, string viewName)
        {
            Document doc = wall.Document;

            if (new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Any(x => x.Name == viewName))
            {
                return new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(x => x.Name == viewName).First();
            }


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

            XYZ xView = - wall.Orientation; //!!!!!!!!!!!!!!!!! And to multiply ot by parameter (-1) изнути 1 снаружи!!!
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

        private static ViewSection CreateHorizontalSectionView(Wall wall, XYZ originPoint, bool upDown, string viewName)
        {

            double look = 1;
            if (upDown)
            { look = 1; }
            else
            { look = -1; }
            Document doc = wall.Document;

            if (new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Any(x => x.Name == viewName))
            {
                return new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(x => x.Name == viewName).First();
            }

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

        public static List<Viewport> GetViewports(View view)
        {
            var collector = new FilteredElementCollector(view.Document)
                .OfClass(typeof(Viewport))
                .WhereElementIsNotElementType()
                .Cast<Viewport>();

            var viewports = collector.Where(vp => vp.ViewId == view.Id).ToList();
            return viewports;
        }
        public static XYZ GetBoundingBoxCenter(Element elem)
        {
            BoundingBoxXYZ boundingBox = elem.get_BoundingBox(null);
            XYZ min = boundingBox.Min;
            XYZ max = boundingBox.Max;
            XYZ center = (min + max) / 2.0;
            return center;
        }
        public static void SetViewportType(View view, string viewportTypeName, Document doc)
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
                        catch (Exception ex)
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

        public static void ShowException(string errorName, Exception ex)
        {
            TaskDialog taskDialog = new TaskDialog("Oh-oh, " + errorName);
            taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
            taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
            taskDialog.Show();
        }
    }
}


