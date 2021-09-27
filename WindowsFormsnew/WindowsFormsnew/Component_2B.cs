using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace WindowsFormsnew
{
    class Component_2B
    {
        public Component_2B(Inventor.Application InventorApplication, double[] Narr, double[] farr, string CB)
        {
            PartDocument oPartdoc;
            oPartdoc = (PartDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject), true);

            PartComponentDefinition oPartCompDef;
            oPartCompDef = oPartdoc.ComponentDefinition;
            PlanarSketch oSketch;
            // PlanarSketch oSketch2;
            oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[3]);


            TransientGeometry oTransGeom = default(TransientGeometry);
            oTransGeom = InventorApplication.TransientGeometry;
            double lengthPS, NR, lengthN, deltay;
            lengthN = farr[6];
            //lengthP = textboxes[2];
            if (CB == "NO")
            {
                lengthPS = Narr[3];
                //lengthPN = Narr[3];
                
            }
            else
            {
                lengthPS = Narr[3];
                //lengthPN = Narr[3] - Farr[6];

            }
            //double NR;
            NR = farr[7] / 2;
            deltay = 0.0625 * 2.54;
            double ofs;
            ofs = 0.625*2.54;
            //double NR;

            Point2d oCoord1, oCoord2;
            oCoord1 = oTransGeom.CreatePoint2d(-lengthPS-ofs, NR);
            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS+farr[10]-ofs, NR);
            SketchLine[] oLine1 = new SketchLine[10];
            oLine1[1] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oCoord1, oCoord2);
            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS+farr[10]-ofs, farr[5]/2);
            oLine1[2] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[1].EndSketchPoint, oCoord2);

            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS+farr[3]-ofs, farr[5] / 2);
            oLine1[3] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[2].EndSketchPoint, oCoord2);

            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS + farr[3]-ofs, farr[2] / 2);
            oLine1[4] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[3].EndSketchPoint, oCoord2);

            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS-ofs+0.0625*2.54, farr[2] / 2);
            oLine1[5] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[4].EndSketchPoint, oCoord2);

            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS - ofs + 0.0625 * 2.54, farr[4] / 2);
            oLine1[6] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[5].EndSketchPoint, oCoord2);

            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS-ofs, farr[4] / 2);
            oLine1[7] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[6].EndSketchPoint, oCoord2);
            oLine1[8] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[1].StartSketchPoint, oLine1[7].EndSketchPoint);
            Profile oProfile = default(Profile);
            oProfile = oSketch.Profiles.AddForSolid();
            RevolveFeature oRevolve;
            WorkAxis oAxis;
            oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[3]);
            oAxis = (WorkAxis)oPartCompDef.WorkAxes[1];
            oRevolve = oPartCompDef.Features.RevolveFeatures.AddByAngle(oProfile, oAxis, 2 * Math.PI, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);

            //functions oextrude = new functions();
            //oextrude.extrude(oPartCompDef, oProfile, 3, 1, 1);
            // for newsolid choose j=2,join choose j=1,else cut
            // for positive direction choose i=1,negative direction choose i=2,else i=3 symmetric
            
            oPartdoc.SaveAs("C:\\Rahul\\Nozzle\\C2.ipt", false);

        }
    }
}
