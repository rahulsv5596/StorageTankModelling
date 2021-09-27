using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace WindowsFormsnew
{
    class Component_3B2
    {
        public Component_3B2(Inventor.Application InventorApplication, double[] Narr, double[] Farr, double[] textboxes)
        {
            PartDocument oPartdoc;
            oPartdoc = (PartDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject), true);

            PartComponentDefinition oPartCompDef;
            oPartCompDef = oPartdoc.ComponentDefinition;
            PlanarSketch oSketch;
            PlanarSketch oSketch2, oSketch3, oSketch4;
            oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[1]);

            TransientGeometry oTransGeom = default(TransientGeometry);
            oTransGeom = InventorApplication.TransientGeometry;
            double lengthP, NR, lengthN, lengthA, WidthA, thickness, radius;
            //lengthN = arr[6];
            //lengthP = textboxes[2];
            lengthA = textboxes[3];
            radius = Narr[2] / 2;
            WidthA = Narr[2];
            thickness = textboxes[2];
            double C_dist;
            // C_dist = Math.Sqrt(radius * radius - Math.Pow(WidthA / 2, 2));
            C_dist = 0;
            //double NR;
            NR = Farr[8] / 2;
            Point2d oCoord1, oCoord2, oCoord3;
            oCoord1 = oTransGeom.CreatePoint2d(-lengthA, -WidthA / 2);
            oCoord2 = oTransGeom.CreatePoint2d(-lengthA, WidthA / 2);
            SketchLine[] oLine1 = new SketchLine[10];
            SketchArc oArc1;

            oLine1[1] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oCoord1, oCoord2);

            oCoord2 = oTransGeom.CreatePoint2d(0, WidthA / 2);

            oLine1[2] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[1].EndSketchPoint, oCoord2);
            oCoord2 = oTransGeom.CreatePoint2d(-1 * C_dist, 0);
            oCoord3 = oTransGeom.CreatePoint2d(0, -WidthA / 2);
            oArc1 = (SketchArc)oSketch.SketchArcs.AddByCenterStartEndPoint(oCoord2, oCoord3, oLine1[2].EndSketchPoint);

            //oCoord2 = oTransGeom.CreatePoint2d(-WidthA / 2, -lengthA);
            oLine1[3] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oArc1.StartSketchPoint, oLine1[1].StartSketchPoint);

            Profile oProfile = default(Profile);
            functions oextrude = new functions();
            oProfile = oSketch.Profiles.AddForSolid();
            oextrude.extrude(oPartCompDef, oProfile, thickness, 1, 2);
            oSketch2 = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[1]);

            SketchCircle oCircle;
            oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            oCircle = (SketchCircle)oSketch2.SketchCircles.AddByCenterRadius(oCoord2, Farr[7] / 2 + (0.125) * 2.54);
            oProfile = oSketch2.Profiles.AddForSolid();
            oextrude.extrude(oPartCompDef, oProfile, thickness, 3, 2);

            //functions custom = new functions();
            //custom.WriteCustomData("Height of Nozzle", textboxes[3], oPartdoc);
            //double outerradius = Farr[7] / 2 + (0.4375) * 2.54;
            //custom.WriteCustomData("Outer Radius", outerradius, oPartdoc);


            oPartdoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
            WorkPlane oWorkPlaneB;
            oWorkPlaneB = (WorkPlane)oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[1], -thickness);
            oSketch3 = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkPlaneB);
            SketchLine oBendLine;
            oCoord1 = oTransGeom.CreatePoint2d(-lengthA, 0);
            oCoord2 = oTransGeom.CreatePoint2d(radius, 0);
            oBendLine = oSketch3.SketchLines.AddByTwoPoints(oCoord1, oCoord2);

            BendDefinition oBenddef;
            ObjectCollection oCol;
            BendPartFeature oBendF;
            oCol = (ObjectCollection)InventorApplication.TransientObjects.CreateObjectCollection();
            oCol.Add(oBendLine);
            //oCol.Add(oProfile2);
            oBendF = oPartCompDef.Features.BendPartFeatures.Add(oBendLine, BendPartTypeEnum.kRadiusAndAngleBendPart, textboxes[1] + textboxes[2], 1.57, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, true);

            ///--> Getting the flat Pattern for the Component code
            SheetMetalComponentDefinition cd;
            FlatPattern fp;
            cd = (SheetMetalComponentDefinition)oPartdoc.ComponentDefinition;
            cd.UseSheetMetalStyleThickness = false;
            cd.Thickness.Value = thickness;
            cd.Unfold();
            cd.FlatPattern.ExitEdit();
            InventorApplication.SilentOperation = true;
            cd.FlatPattern.Body.DataIO.WriteDataToFile("ACIS SAT", "C:\\Rahul\\Nozzle\\C3_1.sat");
            InventorApplication.SilentOperation = false;
            ///
            //FlatPattern oflattpattern;
            //oPartCompDef = oPartdoc.ComponentDefinition;

            functions custom = new functions();
            custom.WriteCustomData("Height of Nozzle", textboxes[3], oPartdoc);
            double outerradius = Farr[7] / 2 + (0.4375) * 2.54;
            custom.WriteCustomData("Outer Radius", outerradius, oPartdoc);

            oPartdoc.SaveAs("C:\\Rahul\\Nozzle\\C3.ipt", false);
            functions oimport = new functions();
            oimport.satimport(InventorApplication, "C:\\Rahul\\Nozzle\\C3_1.sat");
            //TransientObjects transientobj;
            //transientobj = InventorApplication.TransientObjects;
            //NameValueMap option;
            //option = transientobj.CreateNameValueMap();
            //option.Value["ImportUnit"] = 1;



            //InventorApplication.SilentOperation = true;
            oPartdoc = (PartDocument)InventorApplication.ActiveDocument;

            ////InventorApplication.SilentOperation = false;
            oPartCompDef = oPartdoc.ComponentDefinition;

            ObjectCollection oBodies;
            oBodies = InventorApplication.TransientObjects.CreateObjectCollection();
            oBodies.Add(oPartCompDef.SurfaceBodies[1]);
            //oPartCompDef.SurfaceBodies.Count;

            MoveDefinition omoveDef;
            omoveDef = oPartCompDef.Features.MoveFeatures.CreateMoveDefinition(oBodies);

            FreeDragMoveOperation ofreedrag;
            ofreedrag = omoveDef.AddFreeDrag(-WidthA / 2, lengthA, 0);

            RotateAboutLineMoveOperation oRotateAboutAxis;
            //RotateAboutLineMoveOperation oRotateAboutAxis;
            oRotateAboutAxis = omoveDef.AddRotateAboutAxis(oPartCompDef.WorkAxes[2], true, Math.PI / 2);


            oRotateAboutAxis = omoveDef.AddRotateAboutAxis(oPartCompDef.WorkAxes[1], true, Math.PI);



            MoveFeature oMoveFeature;
            oMoveFeature = oPartCompDef.Features.MoveFeatures.Add(omoveDef);

            oPartdoc.Save();


        }
    }
}
