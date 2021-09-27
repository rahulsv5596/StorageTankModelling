using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Windows.Forms;
namespace AssemblyModel
{
    class Hole
    {
        Inventor.Application InventorApplication;
        public void CreateHole(string f, Inventor.Application InventorApplication, AssemblyDocument oAssyDoc,double iangle)
        {
            Inventor.ObjectCollection Bodies;
            Bodies= InventorApplication.TransientObjects.CreateObjectCollection();
            PartDocument oPartDoc;
            oPartDoc=(PartDocument)InventorApplication.Documents.Open(f + "shell\\CBA.ipt", true);
            PartComponentDefinition oPartComp;
            oPartComp = (PartComponentDefinition)oPartDoc.ComponentDefinition;
            TransientGeometry oTransGeom;
            oTransGeom = InventorApplication.TransientGeometry;
          
            //Document odoc = (Document)oPartDoc;
            functions custom = new functions();
            double thickness=Convert.ToDouble(custom.ReadCustomData("Thickness", f + "shell\\CBA.ipt", InventorApplication));
            double N = Convert.ToDouble(custom.ReadCustomData("N", f + "shell\\CBA.ipt", InventorApplication));
            double width = Convert.ToDouble(custom.ReadCustomData("width", f + "shell\\CBA.ipt", InventorApplication));
            double Height = Convert.ToDouble(custom.ReadCustomData("Height of Nozzle", f + "Nozzle\\C3.ipt", InventorApplication));
            double OuterRadius = Convert.ToDouble(custom.ReadCustomData("Outer Radius", f + "Nozzle\\C3.ipt", InventorApplication));



            Double Pi;
            Pi = 3.1415926;
            SurfaceBody oBody;
            oBody = oPartDoc.ComponentDefinition.SurfaceBodies[1];
            Face oface;
            double Radius = 0;
            double angle = 1;
            oface = Maximumface(oBody);
            Arc3d maxoArc=default(Arc3d);
            //Edge oedge;
            foreach(Edge oedge in oface.Edges)
            {

                if (oedge.CurveType == CurveTypeEnum.kLineCurve)
                {
                    Inventor.LineSegment oLine;
                    oLine = oedge.Geometry;
                    
                }
               
                else if(oedge.Geometry is Arc3d)
                {
                    Inventor.Arc3d oArc;
                    oArc = oedge.Geometry;
                    
                    if (oArc.Radius > Radius)
                    {
                        
                        //Radius = oArc.Radius*0.394;//0.394 factor for convertion from cm to in
                        Radius = oArc.Radius;
                        angle = oArc.SweepAngle;
                        maxoArc = oArc;
                        
                    }

                }
                

            }
            
            Point midpoint;
            double dt = 0.125 * 2.54;

            //midpoint = (Point)maxoArc.GetArcData(;
          //  maxoArc.Evaluator.

            //int N;
           // N =(int) ((2 * Pi ) / angle);

            Double angleshift;
            angleshift = (2 * Pi) / (3 * N);
           // dt = dt / (2 * radius);
            Double alfa;
            //R2 = radius;
            alfa = Pi / 2 - 3 * angleshift / 2;
            
            Console.WriteLine("N:{0},Radius:{1}", N,Radius);


             Point occord,occord2;
            WorkPoint oworkpoint1,oworkpoint2;
            int z = 2;


            if (iangle != 0)
            {
                z = (int)(Math.Abs(((iangle + (angle/2)) / angle)))+2;
                
                if (z > N)
                {
                    z = z % (int)N;
                }
              //  z = (int)N - z + 1;
            }
                
                oBody = oPartDoc.ComponentDefinition.SurfaceBodies[z];
            //oBody.
            for(int i = 1; i <= oPartDoc.ComponentDefinition.SurfaceBodies.Count; i++)
            {
                if (i != z)
                {
                    oPartDoc.ComponentDefinition.SurfaceBodies[i].Visible = false;
                }
            }
                oface = Maximumface(oBody);

            // occord = oTransGeom.CreatePoint(Radius * Math.Cos(Pi / N + alfa), 0, Radius * Math.Cos(Pi / N + alfa));

            //occord =(Point) oTransGeom.CreatePoint(Radius * Math.Cos(Pi / N + alfa-iangle), -1*width*2.54, Radius * Math.Sin(Pi / N + alfa-iangle));
            //occord2 = (Point)oTransGeom.CreatePoint(Radius * Math.Cos(Pi / N + alfa -iangle), -1 * width * 2.54+Height, Radius * Math.Sin(Pi / N + alfa-iangle));

            occord = (Point)oTransGeom.CreatePoint(Radius * Math.Cos(Pi / N + alfa - angle - iangle), -1 * width * 2.54, Radius * Math.Sin(Pi / N + alfa - angle - iangle));
            occord2 = (Point)oTransGeom.CreatePoint(Radius * Math.Cos(Pi / N + alfa - angle - iangle), -1 * width * 2.54 + Height, Radius * Math.Sin(Pi / N + alfa - angle - iangle));
            // Point2d occord1;
            // occord1 = oTransGeom.CreatePoint2d(Radius * Math.Cos(Pi / N + alfa), Radius * Math.Cos(Pi / N + alfa));
            functions oextrude = new functions();
            Profile oProfile;
            PartComponentDefinition oPartCompDef;
            oPartCompDef = oPartDoc.ComponentDefinition;
            PlanarSketch oSketch;
            PlanarSketch oSketch2, oSketch3, oSketch4;
            WorkPlane oWorkplane;
            oworkpoint1 = (WorkPoint)oPartComp.WorkPoints.AddFixed(occord);
            
           
            oworkpoint2 = (WorkPoint)oPartComp.WorkPoints.AddFixed(occord2);
            WorkAxis oWorkaxis;
            WorkAxis oWorkaxis2;
            
            oWorkplane = (WorkPlane)oPartComp.WorkPlanes.AddByPointAndTangent(oworkpoint1, oface);
            oWorkaxis = oPartComp.WorkAxes.AddByNormalToSurface(oWorkplane, oworkpoint2);
            oWorkaxis2 = oPartComp.WorkAxes.AddByTwoPoints(oworkpoint1, oworkpoint2);
            Point2d CPoint;
            //MessageBox.Show(oBody.Name);

            if (Math.Abs(iangle) >= Math.Abs(Pi-angle/2) && Math.Abs(iangle) <= Math.Abs(2*Pi-angle/2))
            {
                CPoint = oTransGeom.CreatePoint2d(-1 * Height, 0);
            }
            else
            {
                CPoint = oTransGeom.CreatePoint2d(1 * Height, 0);

            }

            //oPartComp.WorkPlanes.
            //oWorkplane = (WorkPlane)oPartComp.WorkPlanes.AddByPlaneAndOffset(oPartComp.WorkPlanes[2], 100);
            oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkplane);

                SketchCircle oCircle;
                oCircle = (SketchCircle)oSketch.SketchCircles.AddByCenterRadius(CPoint, OuterRadius);
                oProfile = oSketch.Profiles.AddForSolid();

                //Bodies.Add(oBody);
                oextrude.extrude( oPartCompDef, oProfile, Radius / 2, 3, 3);
            //Bodies.Clear();


            //catch (Exception e)
            //{
            //    CPoint = oTransGeom.CreatePoint2d(-1 * Height, 0);

            //    //oPartComp.WorkPlanes.
            //    //oWorkplane = (WorkPlane)oPartComp.WorkPlanes.AddByPlaneAndOffset(oPartComp.WorkPlanes[2], 100);
            //    oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkplane);

            //    SketchCircle oCircle;
            //    oCircle = (SketchCircle)oSketch.SketchCircles.AddByCenterRadius(CPoint, OuterRadius);
            //    oProfile = oSketch.Profiles.AddForSolid();
            //    Bodies.Add(oBody);
            //    oextrude.extrude(oPartCompDef, oProfile, Radius / 2, 3, 3,Bodies);
            //    //Bodies.Clear();

            //}


            for (int i = 1; i <= oPartDoc.ComponentDefinition.SurfaceBodies.Count; i++)
            {
                if (i != z)
                {
                    oPartDoc.ComponentDefinition.SurfaceBodies[i].Visible = true;
                }
            }
            //Bodies.Clear();
            object[] A = new object[3];
            A[0] = oWorkplane;
            A[1] = oWorkaxis;
            A[2] = oWorkaxis2;
            oPartDoc.Save();

            AssemblyComponentDefinition oAssymComp;
            oAssymComp = (AssemblyComponentDefinition)oAssyDoc.ComponentDefinition;
            ComponentOccurrence oC2;
            oC2 = (ComponentOccurrence)oAssymComp.Occurrences[1];
            AssembleHole obj = new AssembleHole(InventorApplication, oAssyDoc,A,oC2);
         
        }


        public Face Maximumface(SurfaceBody oBody)
        {
            
            Double AC = 0;
         

            //oBody = partdoc.ComponentDefinition.SurfaceBodies[1];
            SurfaceEvaluator oFaceEvaluator;
            Face maxFace=null;
            foreach (Face oface in oBody.Faces)
            {
                oFaceEvaluator = oface.Evaluator;
                if (oFaceEvaluator.Area > AC)
                {
                    AC = oFaceEvaluator.Area;
                    maxFace = oface;
                }

            }

            Console.WriteLine(maxFace);

            return maxFace;

        }


    }
}
