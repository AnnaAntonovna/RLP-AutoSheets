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

namespace RLP
{
    [Transaction(TransactionMode.Manual)]
    public class Annotations : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            doc = uiDoc.Document;

            View activeView = doc.ActiveView;

            if (activeView.ViewType != ViewType.Section && activeView.ViewType != ViewType.FloorPlan && activeView.ViewType != ViewType.CeilingPlan)
            {
                // Prompt the user to select a section or plan view
                TaskDialog.Show("View Selection", $"Текущий активный вид имеет тип {activeView.ViewType.ToString()}. Пожалуйста, откройте или активируйте вид типа 'Разрез' или 'План' для проставления марок.");
                return Result.Failed;
            }

            TaskDialogResult result = TaskDialog.Show("Выбор", "Выберите сборку", TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);


            // Get an assembly instance from the user
            Reference assemblyRef = uiDoc.Selection.PickObject(ObjectType.Element, new AssemblySelectionFilter(), "Выберите сборку");

            //var window = new BlockingWindow();
            //window.Show("Processing...");

            AssemblyInstance assemblyInstance = doc.GetElement(assemblyRef) as AssemblyInstance;

            View view = doc.ActiveView;

            // Get all the elements visible in the view
            IList<ElementId> collector = (IList<ElementId>)new FilteredElementCollector(view.Document, view.Id).ToElementIds();
            IList<ElementId> subElementIds = (IList<ElementId>)assemblyInstance.GetMemberIds();

            IList<ElementId> commonElements = collector.Intersect(subElementIds).ToList();

            FilteredElementCollector namecollector = new FilteredElementCollector(doc);
            List<FamilyInstance> familyInstances1 = collector
                .OfType<FamilyInstance>()
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Name.Contains("RLP_Каркас"))
                .ToList();
            List<FamilyInstance> familyInstances2 = collector
                .OfType<FamilyInstance>()
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Name.Contains("IFC_МН"))
                .ToList();

            List<FamilyInstance> familyInstancesCommon = familyInstances1.Concat(familyInstances2).ToList();

            List<ElementId> supercomps = new List<ElementId>(); 

            foreach (FamilyInstance f in familyInstancesCommon)
            {
                if (f.SuperComponent != null)
                {
                    supercomps.Concat(f.GetSubComponentIds());
                }
            }

            IList<ElementId> commonElements2 = commonElements.Except(supercomps).ToList();

            IList <ElementId> windows = (IList<ElementId>)new FilteredElementCollector(view.Document, view.Id)
                .OfCategory(BuiltInCategory.OST_Windows)
                .ToElementIds();
            IList<ElementId> doors = (IList<ElementId>)new FilteredElementCollector(view.Document, view.Id)
                .OfCategory(BuiltInCategory.OST_Doors)
                .ToElementIds();
            IList<ElementId> walls = (IList<ElementId>)new FilteredElementCollector(view.Document, view.Id)
                .OfCategory(BuiltInCategory.OST_Walls)
                .ToElementIds();
            IList<ElementId> parts = (IList<ElementId>)new FilteredElementCollector(view.Document, view.Id)
                .OfCategory(BuiltInCategory.OST_Parts)
                .ToElementIds();

            IList<ElementId> notWindows = commonElements2.Except(windows).ToList();
            IList<ElementId> notParts = notWindows.Except(parts).ToList();
            IList<ElementId> notwalls = notParts.Except(walls).ToList();
            IList<ElementId> finalElements = notParts.Except(doors).ToList();

            using (Transaction tx = new Transaction(doc, "SetTags"))
            {
                tx.Start();

                // Loop through the elements and tag each one on its center
                foreach (ElementId elementId in finalElements)
                {
                    Element element = doc.GetElement(elementId);
                    try
                    {
                        // Get the center point of the element
                        BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                        XYZ center = (bbox.Min + bbox.Max) / 2.0;


                        XYZ centerv = (view.get_BoundingBox(view).Min + view.get_BoundingBox(view).Max) / 2;


                        // Create a new tag
                        IndependentTag tag = IndependentTag.Create(view.Document, view.Id, new Reference(element), false, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, center);

                        // Set the tag leader to point to the center of the element
                        XYZ leaderEnd = center; //+ new XYZ(10, 10, 1);

                        tag.HasLeader = true;
                        XYZ centerView = new XYZ((-view.Outline.Max + view.Outline.Min).U / 2, (-view.Outline.Max + view.Outline.Min).V / 2, 0);


                        //XYZ centerView = (bboxView.Min + bboxView.Max) / 2.0;

                        XYZ vector1 = new XYZ(-2, 10, 0.7);
                        XYZ vector2 = new XYZ(2, 10, 0.7);

                        // Get the location point of the tag
                        // Choose the vector to use for adjusting the location
                        XYZ adjustVector = (center.X < -25) ? vector1 : vector2;

                        tag.Location.Move(adjustVector);

                        tag.LeaderEndCondition = LeaderEndCondition.Free;
                    }
                    catch (InvalidOperationException)
                    {
                        TaskDialog.Show("Не найден тег", $"Для элементов категории {element.Category.Name} не загружена марка аннотаций.");
                    }

                    catch (ArgumentException) { }
                    catch (Exception ex) { }
                    //{
                    //    TaskDialog taskDialog = new TaskDialog("Oh-oh, Can't tag");
                    //    taskDialog.MainContent = ex.Message + '\n' + ex.HelpLink + '\n' + ex.InnerException + '\n' + ex.Data + ex.TargetSite;
                    //    taskDialog.MainInstruction = "...";
                    //    taskDialog.Show();
                    //}
                }
                tx.Commit();
            }
            return Result.Succeeded;
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
}
