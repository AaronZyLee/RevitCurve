using Autodesk.Revit.DB;
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

            int fileNum = 0;

            string path;
            while (true)
            {
                path = "C:\\Users\\DELL\\Documents\\test" + fileNum.ToString() + ".rft";
                if (!File.Exists(path))
                {
                    massdoc.SaveAs(path);
                    break;
                }
                fileNum++;
            }

            string filename = "test" + fileNum.ToString();

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
            XYZ w1_left = LocatePointOnFamilyInstance(well1, Direction.left);
            XYZ w1_right = LocatePointOnFamilyInstance(well1, Direction.right);
            XYZ w2_left = LocatePointOnFamilyInstance(well2, Direction.left);
            XYZ w2_right = LocatePointOnFamilyInstance(well2, Direction.right);

            IList<double> distances = new List<double>();
            distances.Add(w1_left.DistanceTo(w2_left)); 
            distances.Add(w1_left.DistanceTo(w2_right)); 
            distances.Add(w1_right.DistanceTo(w2_left)); 
            distances.Add(w1_right.DistanceTo(w2_right));

            switch(findMinDistance(distances)){
                case(0):
                    CreateCurve(LocatePointOnFamilyInstance(well1, Direction.left), LocatePointOnFamilyInstance(well2, Direction.left), -well1.HandOrientation, -well2.HandOrientation);
                    break;
                case(1):
                    CreateCurve(LocatePointOnFamilyInstance(well1, Direction.left), LocatePointOnFamilyInstance(well2, Direction.right), -well1.HandOrientation, well2.HandOrientation);
                    break;
                case(2):
                    CreateCurve(LocatePointOnFamilyInstance(well1, Direction.right), LocatePointOnFamilyInstance(well2, Direction.left), well1.HandOrientation, -well2.HandOrientation);
                    break;
                case(3):
                    CreateCurve(LocatePointOnFamilyInstance(well1, Direction.right), LocatePointOnFamilyInstance(well2, Direction.right), well1.HandOrientation, well2.HandOrientation);
                    break;
            }

            //CreateCurve(LocatePointOnFamilyInstance(well1, Direction.right), LocatePointOnFamilyInstance(well2, Direction.left), well1.HandOrientation, -well2.HandOrientation);
        }

        private int findMinDistance(IList<double> distances)
        {
            double min = distances[0];
            for(int i=1;i<distances.Count;i++){
                if (distances[i] < min)
                    min = distances[i];
                }
            return distances.IndexOf(min);        

        }

        #region Helper Function

        //绘制模型线
        private void CreateCurve(XYZ startPoint, XYZ endPoint, XYZ normal1, XYZ normal2)
        {
            XYZ StartToEnd = new XYZ((endPoint - startPoint).X, (endPoint - startPoint).Y, 0);
            XYZ p_normal1 = new XYZ(normal1.X, normal1.Y, 0);
            XYZ p_normal2 = new XYZ(normal2.X, normal2.Y, 0);

            p_normal1 = p_normal1 / (Math.Sqrt(p_normal1.X * p_normal1.X + p_normal1.Y * p_normal1.Y));
            p_normal2 = p_normal2 / (Math.Sqrt(p_normal2.X * p_normal2.X + p_normal2.Y * p_normal2.Y));


            XYZ XoYprj_start = new XYZ(startPoint.X, startPoint.Y, 0);
            XYZ XoYprj_end = new XYZ(endPoint.X, endPoint.Y, 0);
            //在起点、终点间插值，并在z=0平面上绘制NurbSpline曲线
            double[] doubleArray = { 1, 1, 1, 1,1,1};
            IList<XYZ> controlPoints2 = new List<XYZ>();
            controlPoints2.Add(XoYprj_start);
            controlPoints2.Add(XoYprj_start + p_normal1 * mmToFeet(2000));
            controlPoints2.Add(XoYprj_start + p_normal1 * mmToFeet(4000));
            controlPoints2.Add(XoYprj_end + p_normal2 * mmToFeet(4000));
            controlPoints2.Add(XoYprj_end + p_normal2 * mmToFeet(2000));
            controlPoints2.Add(XoYprj_end); 

            Curve nbLine = NurbSpline.Create(controlPoints2, doubleArray);
            

            //提取曲线上的拟合点，并在z轴方向插值拟合原曲线
            IList<XYZ> ptsOnCurve = nbLine.Tessellate();
            
            int ptCount = ptsOnCurve.Count;
            ReferencePointArray ptArr = new ReferencePointArray();
            for (int i = 0; i < ptCount; i++)
            {
                XYZ pt = ptsOnCurve[i];
                if(i<(ptCount-1)/8)
                    ptArr.Append(m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, startPoint.Z)));
                else if(i>7*(ptCount-1)/8)
                    ptArr.Append(m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, endPoint.Z )));
                else
                    ptArr.Append(m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, startPoint.Z + (i-(ptCount-1)/8) * (endPoint.Z - startPoint.Z) / (0.75*(ptCount - 1)))));
                //ReferencePoint p = m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, startPoint.Z + i*(endPoint.Z - startPoint.Z)/ (ptCount - 1)));
                //ptArr.Append(p);
            }

            CurveByPoints curve = m_familyCreator.NewCurveByPoints(ptArr);

            curve.Visible = false;

            //创建放样平面并加入参照数组中
            int step = 8;//取step个点进行拟合
            ReferenceArrayArray refArr = new ReferenceArrayArray();
            for (int i = 0; i <= step; i++)
            {
                int position = i * (ptCount - 1) / step;
                
                
                if (i == 0)
                    refArr.Append(CreatePlaneByPoint1(ptArr.get_Item(position), ptArr.get_Item((i + 1) * (ptCount - 1) / step), (curve.GeometryCurve as HermiteSpline).Tangents[position], (curve.GeometryCurve as HermiteSpline).Tangents[((i + 1) * (ptCount - 1) / step)]));
                else if (i == step)
                    refArr.Append(CreatePlaneByPoint1(ptArr.get_Item(position), ptArr.get_Item((i - 1) * (ptCount - 1) / step), (curve.GeometryCurve as HermiteSpline).Tangents[position], (curve.GeometryCurve as HermiteSpline).Tangents[((i - 1) * (ptCount - 1) / step)]));
                    //refArr.Append(CreatePlaneByPoint1(ptArr.get_Item(position), (curve.GeometryCurve as HermiteSpline).Tangents[position]));
                else
                
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), (curve.GeometryCurve as HermiteSpline).Tangents[position]));

            }

            //创建放样实体
            m_familyCreator.NewLoftForm(true, refArr);
            //m_familyCreator.NewFormByCap(true, refArr.get_Item(0));
           

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

        private ReferenceArray CreatePlaneByPoint1(ReferencePoint refPt, ReferencePoint refrefPt , XYZ normal, XYZ refnormal)
        {
            Plane refPlane = new Plane(refnormal, refrefPt.Position);
            Plane plane = new Plane(normal, refPt.Position);
            double sita = -(Math.PI - plane.XVec.AngleOnPlaneTo(refPlane.XVec, plane.Normal));
            Plane planeEx = new Plane(plane.XVec * Math.Cos(sita) + plane.YVec * Math.Sin(sita), plane.YVec * Math.Cos(sita) - plane.XVec * Math.Sin(sita), plane.Origin);
            Arc circle = Arc.Create(planeEx, mmToFeet(300), 0, 2 * Math.PI);
            ModelCurve modelcurve = m_familyCreator.NewModelCurve(circle, SketchPlane.Create(massdoc, planeEx));
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

            EdgeArray edArr = solid.Edges;
            XYZ planeOrigin = (fcArr.get_Item(0) as PlanarFace).Origin;
            return trans.OfPoint(new XYZ(planeOrigin.X,planeOrigin.Y+edArr.get_Item(0).ApproximateLength/2,planeOrigin.Z-edArr.get_Item(1).ApproximateLength/2));
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
