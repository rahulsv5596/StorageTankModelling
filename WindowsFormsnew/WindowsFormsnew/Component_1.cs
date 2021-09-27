//Internal Pipe
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsnew
{
    class Component_1
    {

        public Component_1(Inventor.Application InventorApplication, double[] Narr,double[] Farr,string CB)
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
            double length,lengthPS, lengthPN, NR,deltay;
            Nozzle myform = new Nozzle();
            //length = Narr[3];
            
            if (CB == "NO")
            {
                lengthPS = Narr[3] - Farr[6];
                lengthPN=Narr[3];
                //Console.WriteLine("here we go");
            }
            else
            {
                lengthPS = Narr[3] - Farr[6];
                lengthPN = Narr[3]-Farr[6];

            }
            length = Math.Abs(lengthPN - lengthPS);
            //double NR;
            NR = Farr[8]/2;
            deltay = 0.0625*2.54;
            Point2d oCoord1,oCoord2;
            oCoord1 = oTransGeom.CreatePoint2d(-lengthPS, NR);
            oCoord2 = oTransGeom.CreatePoint2d(+lengthPN, NR);
            SketchLine[] oLine1=new SketchLine[7];
            oLine1[1]=(SketchLine) oSketch.SketchLines.AddByTwoPoints(oCoord1, oCoord2);
            oCoord2 = oTransGeom.CreatePoint2d(lengthPN, NR + deltay);
            oLine1[2] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[1].EndSketchPoint, oCoord2);
            if (CB == "NO")
            {
                

                oCoord2 = oTransGeom.CreatePoint2d(lengthPN , Farr[7] / 2);
                oLine1[3] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[2].EndSketchPoint, oCoord2);

            }
            else
            {
         

                oCoord2 = oTransGeom.CreatePoint2d(lengthPN - 3 * deltay, Farr[7] / 2);
                oLine1[3] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[2].EndSketchPoint, oCoord2);

            }
        
            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS+3*deltay,Farr[7] / 2);
            oLine1[4] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[3].EndSketchPoint, oCoord2);
            oCoord2 = oTransGeom.CreatePoint2d(-lengthPS, NR+deltay);
            oLine1[5] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[4].EndSketchPoint, oCoord2);
            //oCoord2 = oTransGeom.CreatePoint2d(0, 0.5);
            oLine1[6] = (SketchLine)oSketch.SketchLines.AddByTwoPoints(oLine1[5].EndSketchPoint, oLine1[1].StartSketchPoint);
            Profile oProfile = default(Profile);
            oProfile = oSketch.Profiles.AddForSolid();
            RevolveFeature oRevolve;
            WorkAxis oAxis;
            oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[3]);
            oAxis = (WorkAxis)oPartCompDef.WorkAxes[1];
            oRevolve = oPartCompDef.Features.RevolveFeatures.AddByAngle(oProfile, oAxis, 2*Math.PI, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);

            //functions oextrude = new functions();
            //oextrude.extrude(oPartCompDef, oProfile, 3, 1, 1);
            // for newsolid choose j=2,join choose j=1,else cut
            // for positive direction choose i=1,negative direction choose i=2,else i=3 symmetric
            oPartdoc.SaveAs("C:\\Rahul\\Nozzle\\C1.ipt", false);

            //return length;
        }

    }
}
