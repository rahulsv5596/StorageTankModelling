using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
//using Sin()= Math.Sin;

namespace ShellPlate
{
    class CBA
    {

        //Inventor.Application InventorApplication;
        public double CircularBoundaryArray(Inventor.Application InventorApplication,int Level, double radius, double N, double[] H, Double[] Thickness, double dt,string f)
        {
            Double R1;
            Double R2;
            TransientGeometry oTransGeom;
            oTransGeom = InventorApplication.TransientGeometry;
            PartDocument oPartdoc;
            oPartdoc = (PartDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject), true);
            PlanarSketch oSketch;
            Double Pi;
            Pi =  Math.PI;
            PartComponentDefinition oPartCompDef;
            oPartCompDef = oPartdoc.ComponentDefinition;
            //oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[2]);
           
            ExtrudeFeature oExtrude;
            Double delta;
            delta = dt / N;
            //int i;
            Double angleshift;
            angleshift = (2 * Pi) /( 3 * N);
            dt = dt / (2 * radius);
            Double alfa;
            R2 = radius;
            alfa = Pi / 2 - 3 * angleshift / 2;

            for (int j = 1; j <= Level; j++) 
            {
                R1 = R2 + Thickness[j];
                for (int i = 1; i <= N; i++)
                {

                    oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], H[j - 1]));
                    Point2d oCoord1;
                    oCoord1 = oTransGeom.CreatePoint2d(R1 * Math.Cos(((((i - 1) * 2 * Pi) + dt) / N + (j - 1) * angleshift + alfa)), R1 * Math.Sin(((((i - 1) * 2 * Pi + dt)) / N + (j - 1) * angleshift + alfa)));
                    Point2d oCoord2;
                    oCoord2 = oTransGeom.CreatePoint2d(R1 * Math.Cos((((i * 2 * Pi) - dt) / N + (j - 1) * angleshift + alfa)), R1 * Math.Sin((((i * 2 * Pi) - dt) / N + (j - 1) * angleshift + alfa)));

                    SketchArc[] oArc = new SketchArc[3];
                    oArc[1] = (SketchArc)oSketch.SketchArcs.AddByCenterStartEndPoint(oTransGeom.CreatePoint2d(0, 0), oCoord1, oCoord2);
                    oCoord1 = oTransGeom.CreatePoint2d(R2 * Math.Cos((((i - 1) * 2 * Pi) + dt) / N + (j - 1) * angleshift + alfa), R2 * Math.Sin((((i - 1) * 2 * Pi) + dt) / N + (j - 1) * angleshift + alfa));
                    oCoord2 = oTransGeom.CreatePoint2d(R2 * Math.Cos(((i * 2 * Pi) - dt) / N + (j - 1) * angleshift + alfa), R2 * Math.Sin(((i * 2 * Pi) - dt) / N + (j - 1) * angleshift + alfa));
                    oArc[2] = (SketchArc)oSketch.SketchArcs.AddByCenterStartEndPoint(oTransGeom.CreatePoint2d(0, 0), oCoord1, oCoord2);
                    SketchLine[] oLines = new SketchLine[3];
                    
                    oLines[1] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oArc[1].StartSketchPoint, oArc[2].StartSketchPoint);
                    oLines[2] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oArc[1].EndSketchPoint, oArc[2].EndSketchPoint);
                    //oSketch.GeometricConstraints.AddPerpendicular((SketchEntity)oArc[1], (SketchEntity)oLines[1]);
                    //oSketch.GeometricConstraints.AddPerpendicular((SketchEntity)oArc[2], (SketchEntity)oLines[2]);
                    Profile oProfile = default(Profile);
                    oProfile = oSketch.Profiles.AddForSolid();

                    if (Level == 1){



                        //oExtrude = oPartCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, 0.1, kPositiveExtentDirection, kNewBodyOperation);
                        function oExtd = new function();
                        oExtd.extrude(oPartCompDef, oProfile, 0.1, 2, 2);
                    }
                     else
                        {
                        function oExtd = new function();
                        oExtd.extrude(oPartCompDef, oProfile, H[j]-H[j-1]-Math.Round(delta,3), 2, 2);
                        //oExtrude = oPartCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, H[j] - H[j - 1] - Math.Round(delta, 3), kPositiveExtentDirection, kNewBodyOperation);
                    }
                }
                Console.WriteLine("CHECK");
         
            }
            function CustomP = new function();
            Document OdOC = (Document)oPartdoc;
            CustomP.WriteCustomData("Radius", radius/30.48, OdOC);
            CustomP.WriteCustomData("Thickness", Thickness[1], OdOC);
            CustomP.WriteCustomData("N", N, OdOC);
            CustomP.WriteCustomData("level", Level, OdOC);
            CustomP.WriteCustomData("Width", (H[1] - H[0] - Math.Round(delta, 3))/2.54, OdOC);

            oPartdoc.SaveAs(f + "CBA.ipt",false);

            return 3 * angleshift - 0.125 / R2;

        }
    }
}
