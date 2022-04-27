using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]

    public class AddHole : IExternalCommand
    {        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovkDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВиК")).FirstOrDefault();
            if(ovkDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВиК файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("Этаж 1_0.000"))
                .Where(x => x.FamilyName.Equals("ADSK_ОбобщеннаяМодель_ОтверстиеПрямоугольное_вСтене"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство отверстия");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovkDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovkDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transaction0 = new Transaction(arDoc);
            transaction0.Start("Расстановка отверстий");
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            transaction0.Commit();

            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");            
            foreach(Duct d in ducts)
            {
                Line line = (d.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("ADSK_Отверстие_Ширина");
                    Parameter height = hole.LookupParameter("ADSK_Отверстие_Высота");
                    Parameter markFromFloor = hole.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                    width.Set(d.Diameter);
                    height.Set(d.Diameter);
                    double mff = markFromFloor.AsDouble()-d.Diameter/2;                    
                    markFromFloor.Set(mff);
                }
            }
            foreach (Pipe p in pipes)
            {
                Line line = (p.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("ADSK_Отверстие_Ширина");
                    Parameter height = hole.LookupParameter("ADSK_Отверстие_Высота");
                    Parameter markFromFloor = hole.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                    width.Set(p.Diameter);
                    height.Set(p.Diameter);
                    double mff = markFromFloor.AsDouble() - p.Diameter / 2;
                    markFromFloor.Set(mff);
                }
            }
            transaction.Commit();

            return Result.Succeeded;

        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
