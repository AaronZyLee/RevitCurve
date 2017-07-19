﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace createdianlan
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreatenewFamily : IExternalCommand
    {
        
        Autodesk.Revit.Creation.FamilyItemFactory m_familyCreator;
        Document doc;
        Document massdoc;
        Document welldoc;
        const int leftId = 786799;
        const int rightId = 786814;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication app = commandData.Application;
            doc = app.ActiveUIDocument.Document;
            massdoc = app.Application.NewFamilyDocument(@"C:\ProgramData\Autodesk\RVT 2016\Family Templates\Chinese\概念体量\公制体量.rft");
            m_familyCreator = massdoc.FamilyCreate;

            //选择并初始化工井实例
            Autodesk.Revit.UI.Selection.Selection sel = app.ActiveUIDocument.Selection;
            IList<Reference> familylist = sel.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "请选择工井");
            FamilyInstance well1 = doc.GetElement(familylist[0]) as FamilyInstance;
            FamilyInstance well2 = doc.GetElement(familylist[1]) as FamilyInstance;

            //初始化welldoc
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            if (collector != null)
                collector.OfClass(typeof(Family));
            IList<Element> list = collector.ToElements();
            foreach (Element f in list)
            {
                Family family = f as Family;
                if (family.Name == "直通工井")
                {
                    welldoc = doc.EditFamily(family);
                    break;
                }
            }

            //创建电缆族文件并保存
            Transaction trans = new Transaction(massdoc);
            trans.Start("创建新曲线");
            MakeNewCurve(well1, well2);
            trans.Commit();

            string filename = "test24";
            string path = "C:\\Users\\DELL\\Documents\\" + filename + ".rft";
            if (!File.Exists(path))
                massdoc.SaveAs(path);

            //将电缆插入项目文件中
            trans = new Transaction(doc);
            trans.Start("将曲线插入项目文件");
            doc.LoadFamily(path);
            FamilySymbol fs = getSymbolType(doc, filename);
            fs.Activate();
            FamilyInstance fi = doc.Create.NewFamilyInstance(XYZ.Zero, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            fi.Category.Material = getMaterial(doc, "混凝土");
            trans.Commit();

            return Result.Succeeded;

        }

        public FamilySymbol getSymbolType(Document doc, string name)
        {
            FilteredElementIdIterator workWellItrator = new FilteredElementCollector(doc).OfClass(typeof(Family)).GetElementIdIterator();
            workWellItrator.Reset();
            FamilySymbol getsymbol = null;
            while (workWellItrator.MoveNext())
            {
                Family family = doc.GetElement(workWellItrator.Current) as Family;
                foreach (ElementId id in family.GetFamilySymbolIds())
                {
                    FamilySymbol symbol = doc.GetElement(id) as FamilySymbol;
                    if (symbol.Name == name)
                    {
                        getsymbol = symbol;
                    }
                }
            }
            return getsymbol;

        }

        public Material getMaterial(Document doc, string name)
        {
            FilteredElementCollector elementCollector = new FilteredElementCollector(doc).OfClass(typeof(Material));
            IList<Element> materials = elementCollector.ToElements();
            Material getsymbol = null;


            foreach (Element elem in materials)
            {
                Material me = elem as Material;
                if (me.Name.Contains(name))
                {
                    getsymbol = me;
                    if (getsymbol != null)
                        break;
                }
            }
            return getsymbol;

        }

        private void MakeNewCurve(FamilyInstance well1, FamilyInstance well2)
        {
            CreateCurve(LocatePointOnFamilyInstance(well1, Direction.right), LocatePointOnFamilyInstance(well2, Direction.left), well1.HandOrientation, -well2.HandOrientation);
        }

        #region Helper Function

        //绘制模型线
        private void CreateCurve(XYZ startPoint, XYZ endPoint, XYZ normal1, XYZ normal2)
        {
            XYZ StartToEnd = new XYZ((endPoint - startPoint).X, (endPoint - startPoint).Y, 0);
            XYZ p_normal1 = new XYZ(normal1.X, normal1.Y, 0);
            XYZ p_normal2 = new XYZ(normal2.X, normal2.Y, 0);

            /**
            double sita1 = StartToEnd.AngleTo(p_normal1);
            double sita2 = StartToEnd.AngleTo(p_normal2);
            XYZ e1 = Math.Abs(((0.25 * StartToEnd).GetLength() * Math.Cos(sita1) / p_normal1.GetLength())) * p_normal1;
            XYZ e2 = Math.Abs(((0.25 * StartToEnd).GetLength() * Math.Cos(sita2) / p_normal2.GetLength())) * p_normal2;
            

            XYZ prePt1 = new XYZ(startPoint.X + e1.X, startPoint.Y + e1.Y, 0);
            XYZ prePt2 = new XYZ(endPoint.X + e2.X, endPoint.Y + e2.Y, 0);
             */
            p_normal1 = p_normal1 / (Math.Sqrt(p_normal1.X * p_normal1.X + p_normal1.Y * p_normal1.Y));
            p_normal2 = p_normal2 / (Math.Sqrt(p_normal2.X * p_normal2.X + p_normal2.Y * p_normal2.Y));     


            //在起点、终点间插值，绘制NurbSpline曲线
            double[] doubleArray = { 1, 1, 1, 1,1,1};
            IList<XYZ> controlPoints2 = new List<XYZ>();

            /**
            controlPoints2.Insert(0, new XYZ(startPoint.X, startPoint.Y, 0));
            controlPoints2.Insert(1, prePt1);
            //controlPoints2.Insert(2, new XYZ((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2, 0));
            controlPoints2.Insert(2, prePt2);
            controlPoints2.Insert(3, new XYZ(endPoint.X, endPoint.Y, 0));
            */

            controlPoints2.Add(startPoint);
            controlPoints2.Add(startPoint + p_normal1 * mmToFeet(2000));
            controlPoints2.Add(startPoint + p_normal1 * mmToFeet(4000));
            controlPoints2.Add(endPoint + p_normal2 * mmToFeet(4000));
            controlPoints2.Add(endPoint + p_normal2 * mmToFeet(2000));
            controlPoints2.Add(endPoint); 

            Curve nbLine = NurbSpline.Create(controlPoints2, doubleArray);



            //提取曲线上的拟合点
            IList<XYZ> ptsOnCurve = nbLine.Tessellate();
            
            int ptCount = ptsOnCurve.Count;
            ReferencePointArray ptArr = new ReferencePointArray();
            for (int i = 0; i < ptCount; i++)
            {
                XYZ pt = ptsOnCurve[i];
                ReferencePoint p = m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, startPoint.Z + i / (ptCount-1) * (endPoint.Z - startPoint.Z)));
                ptArr.Append(p);
            }

            CurveByPoints curve = m_familyCreator.NewCurveByPoints(ptArr);

            curve.Visible = false;

            //创建放样平面并加入参照数组中
            int step = 4;//取4分点进行拟合
            ReferenceArrayArray refArr = new ReferenceArrayArray();
            for (int i = 0; i < step; i++)
            {
                int position = i * (ptCount - 1) / step;
                if (i == 0)
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), normal1));
                else if (i == ptArr.Size - 1)
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), normal2));
                else
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), (curve.GeometryCurve as HermiteSpline).Tangents[position]));

            }

            //创建放样实体
            m_familyCreator.NewLoftForm(true, refArr);

        }

        //根据参照点和法向量创建放样截面
        private ReferenceArray CreatePlaneByPoint(ReferencePoint refPt, XYZ normal)
        {
            Plane plane = new Plane(normal, refPt.Position);
            Arc circle = Arc.Create(plane, mmToFeet(300), 0, 2 * Math.PI);
            ModelCurve modelcurve = m_familyCreator.NewModelCurve(circle, SketchPlane.Create(massdoc, plane));
            ReferenceArray ra = new ReferenceArray();
            ra.Append(modelcurve.GeometryCurve.Reference);
            return ra;
        }

        //在族实例上定位连接点
        private XYZ LocatePointOnFamilyInstance(FamilyInstance fi, Direction dir)
        {
            //找到工井族文件中左边界或右边界的几何体，利用id定位
            Options opt = new Options();
            opt.IncludeNonVisibleObjects = true;//很重要，由于view为null需要将不可见的对象变为可见，否则提取空集
            GeometryElement ge = null;
            switch(dir){
                case(Direction.left):
                    ge = welldoc.GetElement(new ElementId(leftId)).get_Geometry(opt);
                    break;
                case(Direction.right):
                    ge = welldoc.GetElement(new ElementId(rightId)).get_Geometry(opt);
                    break;
            }

            //提取集合体上的面元素
            Solid solid = null;
            foreach (GeometryObject obj in ge)
            {
                if (obj is Solid)
                {
                    solid = obj as Solid;
                    break;
                }
            }
            FaceArray fcArr = null;
            if(solid != null){
                fcArr = solid.Faces;
            }

            //找到边界的面，并将面原点的族坐标转换为模型坐标
            Transform trans = null;
            Options opt1 = new Options();
            opt1.ComputeReferences = false;
            opt1.View = doc.ActiveView;
            GeometryElement geoElement = fi.get_Geometry(opt1);
            foreach (GeometryObject obj in geoElement)   //利用GeometryInstrance的Transform提取坐标转换矩阵
            {
                if (obj is GeometryInstance)
                {
                    GeometryInstance inst = obj as GeometryInstance;
                    trans = inst.Transform;
                    break;
                }
            }
            if(dir == Direction.left)
                return trans.OfPoint((fcArr.get_Item(0) as PlanarFace).Origin);

            return trans.OfPoint((fcArr.get_Item(0) as PlanarFace).Origin);
        }

        //枚举类Direction
        private enum Direction
        {
            left,
            right,
        }

        //根据类型及名称查找元素
        Element findElement(Document doc, Type targetType, string targetName)
        {
            // get the elements of the given type
            //
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.WherePasses(new ElementClassFilter(targetType));

            // parse the collection for the given name
            // using LINQ query here. 
            // 
            var targetElems = from element in collector where element.Name.Equals(targetName) select element;
            List<Element> elems = targetElems.ToList<Element>();

            if (elems.Count > 0)
            {  // we should have only one with the given name. 
                return elems[0];
            }

            // cannot find it.
            return null;
        }

        //毫米英尺单位转化
        double mmToFeet(double mmVal)
        {
            return mmVal / 304.8;
        }

        #endregion

    }
}
