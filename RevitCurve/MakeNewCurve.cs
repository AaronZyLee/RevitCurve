using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Creation;

namespace RevitCurve
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [Journaling(JournalingMode.NoCommandData)]
    public class Command : IExternalCommand{

        private static Autodesk.Revit.ApplicationServices.Application m_application;
        private static Autodesk.Revit.DB.Document m_document;
        private Autodesk.Revit.Creation.Application m_appCreator;
        private FamilyItemFactory m_familyCreator;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                m_application = commandData.Application.Application;
                m_document = commandData.Application.ActiveUIDocument.Document;


                if (!m_document.IsFamilyDocument)
                {
                    message = "无法在非族类文档中使用";
                    return Result.Failed;
                }


                m_appCreator = m_application.Create;
                m_familyCreator = m_document.FamilyCreate;

                Transaction trans = new Transaction(m_document);
                trans.Start("创建新曲线");

                MakeNewCurve();
                
                //commandData.Application.ActiveUIDocument.Selection.PickPoint

                trans.Commit();

            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
            return Result.Succeeded;
        }

        private void MakeNewCurve() {

            /*
            View pViewPlan = findElement(typeof(ViewPlan), "{三维}") as View;
            
            //get the ref plane at the one of the predefined plane
            ReferencePlane ctrLR = findElement(typeof(ReferencePlane), "中心(左/右)") as ReferencePlane;

            XYZ p1 = ctrLR.BubbleEnd;
            XYZ p2 = ctrLR.FreeEnd;
            XYZ p3 = new XYZ(0, 500, 1000);
            Line l1 = Line.CreateBound(p1, p2);
            Line l2 = Line.CreateBound(p1, p3);

            SketchPlane sp = SketchPlane.Create(m_document,ctrLR.GetPlane());

            m_familyCreator.NewModelCurve(l1, sp);
            m_familyCreator.NewModelCurve(l2, sp);
            
            
            #region CurvebyPoints
            IList<XYZ> controlPoints = new List<XYZ>();
            controlPoints.Insert(0,new XYZ(-175,0,-70));
            controlPoints.Insert(1,new XYZ(30,50,104));
            controlPoints.Insert(2,new XYZ(-70,-20,60));
            //controlPoints.Insert(3,new XYZ(-222,100,100));

            ReferencePointArray refArray = new ReferencePointArray();
            for (int i = 0; i < controlPoints.Count; i++)
            {
                ReferencePoint refPt = m_familyCreator.NewReferencePoint(controlPoints[i]);
                refArray.Append(refPt);
            }
            CurveByPoints c1 = m_familyCreator.NewCurveByPoints(refArray);

            ReferenceArrayArray profilesArray = new ReferenceArrayArray();

            for (int i = 0; i < refArray.Size; i++)
            {
                Plane plane = new Plane((c1.GeometryCurve as HermiteSpline).Tangents[i], c1.GetPoints().get_Item(i).Position);
                Arc circle = Arc.Create(plane, mmToFeet(1000), 0, 2 * Math.PI);
                ModelCurve modelcurve = m_familyCreator.NewModelCurve(circle, SketchPlane.Create(m_document, plane));
                ReferenceArray profile = new ReferenceArray();
                profile.Append(modelcurve.GeometryCurve.Reference);
                profilesArray.Append(profile);

            }

            Form form = m_familyCreator.NewLoftForm(true, profilesArray);
            #endregion
            

            #region NurbSpline
            double[] doubleArray = {1,1,1,1};
            IList<XYZ> controlPoints2 = new List<XYZ>();
            controlPoints2.Insert(0,new XYZ(0,10,20));
            controlPoints2.Insert(0,new XYZ(0,20,50));
            controlPoints2.Insert(0,new XYZ(0,30,0));
            controlPoints2.Insert(0,new XYZ(0,50,50));
            Curve nbLine = NurbSpline.Create(controlPoints2, doubleArray);

            

            //提取曲线上的拟合点
            IList<XYZ> ptsOnCurve = nbLine.Tessellate();
            Plane yz = new Plane(new XYZ(0,1,0),new XYZ(0,1,1 ),XYZ.Zero);
            m_familyCreator.NewModelCurve(nbLine,SketchPlane.Create(m_document,yz));

            //m_familyCreator.NewModelCurve(nbLine,SketchPlane.Create(m_document,(findElement(typeof(ReferencePlane),"中心(左/右)") as ReferencePlane).GetPlane()));

            #endregion
            **/


            CreateCurve1(new XYZ(200, 100, 10), new XYZ(-120, -50, 50), new XYZ(-1, 0, 0), new XYZ(0, 1, 0));


            /*Line l1 = Line.CreateBound(new XYZ(100,100,100),new XYZ(500,500,500));
            
            Plane plane = new Plane(m_document.ActiveView.ViewDirection, m_document.ActiveView.Origin);
            SketchPlane sp = SketchPlane.Create(m_document,plane);

            Curve m_curve = Autodesk.Revit.DB.HermiteSpline.Create(controlPoints, false);

            ReferencePointArray ptArray = new ReferencePointArray();
            ReferencePoint p1 = new ReferencePoint();
            p1.Position = new XYZ(-175,0,-70);
            ptArray.Append(p1);
            ReferencePoint p2 = new ReferencePoint();
            p2.Position = new XYZ(30,50,104);
            ptArray.Append(p2); 




            m_familyCreator.NewCurveByPoints(ptArray);

            //m_familyCreator.NewModelCurve(l1,sp);
             **/
        }


        private void CreateCurve1(XYZ startPoint, XYZ endPoint, XYZ normal1, XYZ normal2)
        {
            XYZ StartToEnd = new XYZ((endPoint - startPoint).X, (endPoint - startPoint).Y, 0);
            XYZ p_normal1 = new XYZ(normal1.X, normal1.Y, 0);
            XYZ p_normal2 = new XYZ(normal2.X, normal2.Y, 0);

            p_normal1 = p_normal1 / (Math.Sqrt(p_normal1.X * p_normal1.X + p_normal1.Y * p_normal1.Y));
            p_normal2 = p_normal2 / (Math.Sqrt(p_normal2.X * p_normal2.X + p_normal2.Y * p_normal2.Y));


            XYZ XoYprj_start = new XYZ(startPoint.X, startPoint.Y, 0);
            XYZ XoYprj_end = new XYZ(endPoint.X, endPoint.Y, 0);
            //在起点、终点间插值，并在z=0平面上绘制NurbSpline曲线
            double[] doubleArray = { 1, 1, 1, 1, 1, 1 };
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
                ReferencePoint p = m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, startPoint.Z + i * (endPoint.Z - startPoint.Z) / (ptCount - 1)));
                ptArr.Append(p);
            }

            CurveByPoints curve = m_familyCreator.NewCurveByPoints(ptArr);

            curve.Visible = false;

            //创建放样平面并加入参照数组中
            int step = 8;//取8个点进行拟合
            ReferenceArrayArray refArr = new ReferenceArrayArray();
            for (int i = 0; i <= step; i++)
            {
                int position = i * (ptCount - 1) / step;
                /**
                if (i == 0)
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), normal1));
                else if (i == step)
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), normal2));
                else
                 */
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), (curve.GeometryCurve as HermiteSpline).Tangents[position]));

            }

            //创建放样实体
            m_familyCreator.NewLoftForm(true, refArr);

        }


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
            double[] doubleArray = { 1, 1, 1, 1, 1, 1 };
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
                ReferencePoint p = m_familyCreator.NewReferencePoint(new XYZ(pt.X, pt.Y, startPoint.Z + i / (ptCount - 1) * (endPoint.Z - startPoint.Z)));
                ptArr.Append(p);
            }

            CurveByPoints curve = m_familyCreator.NewCurveByPoints(ptArr);

            curve.Visible = false;

            //创建放样平面并加入参照数组中
            int step = 16;//取4分点进行拟合
            ReferenceArrayArray refArr = new ReferenceArrayArray();
            for (int i = 0; i <= step; i++)
            {
                int position = i * (ptCount - 1) / step;
                if (i == 0)
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), normal1));
                else if (i == ptArr.Size - 1)
                /**
                {
                    Plane plane = new Plane(normal2, ptArr.get_Item(position).Position);
                    Plane planeEx = new Plane(plane.XVec.CrossProduct(normal2),plane.YVec.CrossProduct(normal2),plane.Origin);
                }
                */   
                    refArr.Append(CreatePlaneByPoint(ptArr.get_Item(position), -normal2));
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
            ModelCurve modelcurve = m_familyCreator.NewModelCurve(circle, SketchPlane.Create(m_document, plane));
            ReferenceArray ra = new ReferenceArray();
            ra.Append(modelcurve.GeometryCurve.Reference);
            return ra;
        }


        #region Helper Functions

        // ==================================================================================
        //   helper function: find an element of the given type and the name.
        //   You can use this, for example, to find Reference or Level with the given name.
        // ==================================================================================
        Element findElement(Type targetType, string targetName)
        {
            // get the elements of the given type
            //
            FilteredElementCollector collector = new FilteredElementCollector(m_document);
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

        // ===============================================
        //   helper function: convert millimeter to feet
        // ===============================================
        double mmToFeet(double mmVal)
        {
            return mmVal / 304.8;
        }

        #endregion // Helper Functions
    }
}
