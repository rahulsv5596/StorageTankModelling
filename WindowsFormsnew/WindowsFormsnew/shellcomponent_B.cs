using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
namespace WindowsFormsnew
{
    class shellcomponent_B
    {
        public shellcomponent_B(Inventor.Application InventorApplication, double[] Narr, double[] Farr, double[] textboxes)
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
            Profile oProfile1, oProfile2, oProfile3;
            thickness = textboxes[2];
            //lengthN = arr[6];
            //lengthP = textboxes[2];
            SketchLines oRect = default(SketchLines);
            Point2d oCoord1, oCoord2;
            oCoord1 = oTransGeom.CreatePoint2d(-textboxes[3], -textboxes[5] / 2);
            oCoord2 = oTransGeom.CreatePoint2d(-textboxes[3] + textboxes[6], textboxes[5] / 2);

            oSketch.SketchLines.AddAsTwoPointRectangle(oCoord1, oCoord2);



            oProfile1 = oSketch.Profiles.AddForSolid();
            functions oextrude = new functions();
            oextrude.extrude(oPartCompDef, oProfile1, textboxes[2], 1, 1);


            oSketch2 = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[1]);

            SketchCircle oCircle;
            oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            oCircle = (SketchCircle)oSketch2.SketchCircles.AddByCenterRadius(oCoord2, Farr[7] / 2 + (0.3125) * 2.54);
            oProfile2 = oSketch2.Profiles.AddForSolid();
            oextrude.extrude(oPartCompDef, oProfile2, textboxes[2], 3, 1);
            //WorkPlane oWorkPlane;
            //oWorkPlane = (WorkPlane)oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[1], thickness);
            //oSketch3 = (PlanarSketch)oPartCompDef.Sketches.Add(oWorkPlane);
            //oCoord2 = oTransGeom.CreatePoint2d(0, 0);
            //oCircle = (SketchCircle)oSketch3.SketchCircles.AddByCenterRadius(oCoord2, Farr[7] / 2 + (0.4375) * 2.54);
            //Profile oProfile3;

            //oProfile3 = oSketch3.Profiles.AddForSolid();



            ///Loft Function
            //LoftDefinition oLoftdef;
            //ObjectCollection oCol;
            //oCol = (ObjectCollection)InventorApplication.TransientObjects.CreateObjectCollection();
            //oCol.Add(oProfile2);
            //oCol.Add(oProfile3);
            ////oCol.Add(oProfile3);
            //LoftFeature oLoftF;

            //oLoftdef = oPartCompDef.Features.LoftFeatures.CreateLoftDefinition(oCol, PartFeatureOperationEnum.kCutOperation);
            //oLoftF = oPartCompDef.Features.LoftFeatures.Add(oLoftdef);

            oPartdoc.SaveAs("C:\\Rahul\\Nozzle\\C4.ipt", false);
        }
    }
}
