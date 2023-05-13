/*
 * Created by SharpDevelop.
 * User: an2ba
 * Date: 5/11/2023
 * Time: 11:03 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace MacroHelper
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("198D8183-9331-45E4-90D5-1B1BCD19CDF4")]
	public partial class ThisApplication
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
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
		
		
		
		
//		public void hidehide()
//		{
//			// Get the active document and the user interface document
//		    Document doc = this.ActiveUIDocument.Document;
//		    UIDocument uidoc = this.ActiveUIDocument;
//		    
//			TaskDialogResult result = TaskDialog.Show("Выбор", "Выберите сборку", TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);
//			
//
//            // Get an assembly instance from the user
//            Reference assemblyRef = uidoc.Selection.PickObject(ObjectType.Element, new AssemblySelectionFilter(), "Выберите сборку");
//
//            //var window = new BlockingWindow();
//            //window.Show("Processing...");
//
//            AssemblyInstance assemblyInstance = doc.GetElement(assemblyRef) as AssemblyInstance;
//			
//            View view = doc.ActiveView;
//            Wall hostWall = null;
//            
//           //TaskDialogResult result1 = TaskDialog.Show("Выбор", "Выберите wall", TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);
//
//            // Get an assembly instance from the user
//            //Reference hostWallRef = uidoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Выберите сборку");
//
//            //var window = new BlockingWindow();
//            //window.Show("Processing...");
//
//            //hostWall = doc.GetElement(hostWallRef) as Wall;
//            
//            
//
//                    //METHODS
//			using (Transaction tx = new Transaction(doc, "CreateWallSectionView"))
//                {
//                    tx.Start();
//                   
//            
//            FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id);
//            ICollection<ElementId> elementIdsToHide = collector
//                .Where(e => e.Id != assemblyInstance.Id && e.Id != view.Id && e.CanBeHidden(view)
//                && !assemblyInstance.GetMemberIds().ToList().Contains(e.Id))
//                .Select(e => e.Id)
//                .ToList();
//
//            // Hide all elements except the hostWall and assembly
//            
//            view.HideElements(elementIdsToHide);
//            tx.Commit();}
//			
//		}
		public void SectionView()
		{
			Document doc = this.ActiveUIDocument.Document;
		    UIDocument uidoc = this.ActiveUIDocument;
		    
		    // Ask the user to select a wall
		        Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Please select a wall");
		
		        // Get the wall from the reference
		        Wall wall = doc.GetElement(pickedRef) as Wall;
		        
		        List<XYZ> list = GetPointsForVerticalSections(wall, false);
		        if (list != null && list.Count > 0)
		        {
			        foreach (XYZ point in list)
			        {
			        	//TaskDialog.Show("Points", point.ToString() + "\n" + list.Count.ToString());
			        	using (Transaction tx = new Transaction(doc, "CreateWallSectionView"))
		                {
		                    tx.Start();
			        		ViewSection sv =  CreateVerticalSectionView(wall, point);
			        		tx.Commit();
			        		uidoc.ActiveView = sv;
			        	}
			        }
		        }
		  			
		        
		        List<XYZ> list2 = GetPointsForHorizontalSections(wall, 0);
		        if (list2 != null && list.Count > 0)
		        {
		        	bool up = true; //to set this to a first section method
			        foreach (XYZ point in list2)
			        {
			        	using (Transaction tx = new Transaction(doc, "CreateWallSectionView"))
		                {
		                    tx.Start();
			        	//TaskDialog.Show("Points", point.ToString() + "\n" + list2.Count.ToString());
			        		ViewSection sv =  CreateHorizontalSectionView(wall, point, up);
			        		tx.Commit();
			        		
			        		uidoc.ActiveView = sv;
			        	}
			        	up = false;
			        }
		        }
			}
		
		public List<XYZ> GetPointsForVerticalSections(Wall wall, bool opalVnutr)
        {
			Document doc = this.ActiveUIDocument.Document;
		    UIDocument uidoc = this.ActiveUIDocument;
			
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
                            if ((cPplus - centerPoints[0] ).GetLength() < offsetDistance)
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
            Document doc = this.ActiveUIDocument.Document;
            UIDocument uidoc = this.ActiveUIDocument;

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
            XYZ  viewdir = xView.CrossProduct(yView);
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
         		{look = 1;}
         	else
         		{look = -1;}
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
			
            XYZ  viewdir = XYZ.BasisZ * look;
            
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
        
         
         public void SetParameters(ViewSection view, string assemblyName, string viewName, string viewTemplateName, string viewportTypeName)
        {
            try
            {
                // Set the view name and viewport detail number parameters
                view.get_Parameter(BuiltInParameter.VIEW_NAME).Set(viewName);
                view.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set(assemblyName);

                // Apply the view template
                ApplyViewTemplate(view, viewTemplateName);

            }
            catch (Exception ex)
            {
                TaskDialog taskDialog = new TaskDialog("Oh-oh, can't SetParameters");
                taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
                taskDialog.Show();
            }
        }
         
         public void ShowException(string errorName, Exception ex)
        {
            TaskDialog taskDialog = new TaskDialog("Oh-oh, " + errorName);
            taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
            taskDialog.MainInstruction = "Вы поймали ошибку, которую мы еще не видели! Пожалуйста, поделитесь этим прекрасным открытием с нами! Сообщите о проблеме разработчику и опишите ее по адресу - an2baranova@gmail.com. А пока вы ждете ответа, убедитесь, что действовали по инструкции. Приятного пользования!";
            taskDialog.Show();
        }
         
         
         public class WallSelectionFilter : ISelectionFilter
		{
		    public bool AllowElement(Element elem)
		    {
		        return elem is Wall;
		    }
		
		    public bool AllowReference(Reference reference, XYZ position)
		    {
		        return true;
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
                TaskDialog.Show("Error", "The view template named {viewTemplateName} does not exist.");
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
         
         
        
    }	
	}
