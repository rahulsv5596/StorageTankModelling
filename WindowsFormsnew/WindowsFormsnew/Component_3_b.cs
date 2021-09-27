//middle portions

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace WindowsFormsnew
{
    class Component_3_b
    {
        public Component_3_b(Inventor.Application InventorApplication, double[] Narr, double[] Farr, double[] textboxes)
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
            //double C_dist;
            // C_dist = Math.Sqrt(radius * radius - Math.Pow(WidthA / 2, 2));
            //C_dist = 0;
            //double NR;
            NR = Farr[8] / 2;
            Point2d oCoord2,oCoord1;

            SketchCircle oCircle2;
            oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            oCircle2 = (SketchCircle)oSketch.SketchCircles.AddByCenterRadius(oCoord2, radius);

            Profile oProfile = default(Profile);
            functions oextrude = new functions();
            oProfile = oSketch.Profiles.AddForSolid();
            oextrude.extrude(oPartCompDef, oProfile, thickness, 1, 2);
            oSketch2 = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[1]);

            SketchCircle oCircle;
            oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            oCircle = (SketchCircle)oSketch2.SketchCircles.AddByCenterRadius(oCoord2, Farr[7] / 2 + (0.125) * 2.54);
            oProfile = oSketch2.Profiles.AddForSolid();
            WorkPlane oWorkPlane;
            oWorkPlane = (WorkPlane)oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[1], -1 * thickness);
            oSketch3 = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkPlane);
            oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            oCircle = (SketchCircle)oSketch3.SketchCircles.AddByCenterRadius(oCoord2, Farr[7] / 2 + (0.4375) * 2.54);
            Profile oProfile2;

            oProfile2 = oSketch3.Profiles.AddForSolid();

            WorkPlane oWorkPlane2;
            oWorkPlane2 = (WorkPlane)oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[1], -1 * thickness);
            oSketch4 = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkPlane);
            oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            oCircle = (SketchCircle)oSketch4.SketchCircles.AddByCenterRadius(oCoord2, Farr[7] / 2 + (0.4375) * 2.54);
            // Profile oProfile3;
            //oProfile3 = oSketch4.Profiles.AddForSolid();
            LoftDefinition oLoftdef;
            ObjectCollection oCol;
            oCol = (ObjectCollection)InventorApplication.TransientObjects.CreateObjectCollection();
            oCol.Add(oProfile);
            oCol.Add(oProfile2);
            //oCol.Add(oProfile3);
            LoftFeature oLoftF;
            
            oLoftdef = oPartCompDef.Features.LoftFeatures.CreateLoftDefinition(oCol, PartFeatureOperationEnum.kCutOperation);
            oLoftF = oPartCompDef.Features.LoftFeatures.Add(oLoftdef);

            oPartdoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
            WorkPlane oWorkPlaneB;
            oWorkPlaneB = (WorkPlane)oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[1], -thickness);
            oSketch3 = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkPlaneB);
            SketchLine oBendLine;
            oCoord1 = oTransGeom.CreatePoint2d(-radius, 0);
            oCoord2 = oTransGeom.CreatePoint2d(radius, 0);
            oBendLine = oSketch3.SketchLines.AddByTwoPoints(oCoord1, oCoord2);

            BendDefinition oBenddef;
            ObjectCollection ocol2;
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
            //oextrude.extrude(oPartCompDef, oProfile2, 0.125*2.54, 3, 2);

            functions custom = new functions();
            custom.WriteCustomData("Height of Nozzle", textboxes[3], oPartdoc);
            double outerradius = Farr[7] / 2 + (0.4375) * 2.54;
            custom.WriteCustomData("Outer Radius", outerradius, oPartdoc);
            oPartdoc.SaveAs("C:\\Rahul\\Nozzle\\C3.ipt", false);
            //oPartdoc.SaveAs("C:\\Rahul\\Nozzle\\C3.ipt", false);
            functions oimport = new functions();
            oimport.satimport(InventorApplication, "C:\\Rahul\\Nozzle\\C3_1.sat");

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
            ofreedrag = omoveDef.AddFreeDrag(0, -radius, 0);

            RotateAboutLineMoveOperation oRotateAboutAxis;
            //RotateAboutLineMoveOperation oRotateAboutAxis;
            oRotateAboutAxis = omoveDef.AddRotateAboutAxis(oPartCompDef.WorkAxes[2], true, Math.PI / 2);


            //oRotateAboutAxis = omoveDef.AddRotateAboutAxis(oPartCompDef.WorkAxes[1], true, Math.PI);



            MoveFeature oMoveFeature;
            oMoveFeature = oPartCompDef.Features.MoveFeatures.Add(omoveDef);

            oPartdoc.Save();

        }


    }
}
