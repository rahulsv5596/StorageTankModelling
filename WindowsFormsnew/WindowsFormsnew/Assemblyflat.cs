using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace WindowsFormsnew
{
    class Assemblyflat
    {
        public Assemblyflat(Inventor.Application InventorApplication, double[] arr, double[] textboxes, string CB)
        {
            AssemblyDocument oAssyDoc;
            oAssyDoc = (AssemblyDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);
            TransientGeometry oTransGeom = default(TransientGeometry);
            oTransGeom = InventorApplication.TransientGeometry;

            Matrix oPositionMatrix, oPositionMatrix2;

            oPositionMatrix = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
            oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();

            string sfilename;
            sfilename = "C:\\Rahul\\Nozzle\\C1.ipt";
            ComponentOccurrence oC1;
            oC1 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);

            Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
            oPositionMatrix.SetTranslation(oTrans);


            sfilename = "C:\\Rahul\\Nozzle\\C2.ipt";
            ComponentOccurrence oC2A, oC2B;
            oC2A = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);
            oC2B = default(ComponentOccurrence);
            Point oCoord1;
            oCoord1 = (Point)oTransGeom.CreatePoint(0, 0, 0);
            // Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(0, 1, 0);
            oPositionMatrix.SetTranslation(oTrans);


            if (CB == "YES")
            {
                sfilename = "C:\\Rahul\\Nozzle\\C2.ipt";
                //ComponentOccurrence oC2B;
                oC2B = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix2);

                //Vector oTrans;
                // oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
                oPositionMatrix2.SetToRotation(Math.PI, oTrans, oCoord1);
            }

            sfilename = "C:\\Rahul\\Nozzle\\C3_1.ipt";
            ComponentOccurrence oC3;
            oC3 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);

            // Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
            oPositionMatrix.SetTranslation(oTrans);

            sfilename = "C:\\Rahul\\Nozzle\\C4.ipt";
            ComponentOccurrence oC4;
            oC4 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);

            // Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
            oPositionMatrix.SetTranslation(oTrans);

            //geting face details

            //component 1
            //Face oface;





            //for (k = 1; k < faceC; k++)
            //{
            //    Console.WriteLine("Areas of the faces {0} are {1}",k, Area[k]);
            //}

            // k = 1;

            int k = 1;
            int faceC;

            //component2
            faceC = oC1.SurfaceBodies[1].Faces.Count;
            Face[] oCB = new Face[faceC + 1];
            double[] Area = new double[faceC + 1];
            //WorkAxes oAxis;

            SurfaceEvaluator ofaceEval;


            foreach (Face oface in oC1.SurfaceBodies[1].Faces)
            {
                oCB[k] = oface;
                ofaceEval = oface.Evaluator;
                Area[k] = ofaceEval.Area;
                k++;
            }


            for (k = 1; k < faceC; k++)
            {
                Console.WriteLine("Areas of the faces {0} are {1}", k, Area[k]);
            }






            functions A_C = new functions();
            A_C.FaceContraints(oAssyDoc, oC1, oC2A, 1, 8);
            A_C.AxisContraints(oAssyDoc, oC1, oC2A, 1, 1);

            if (CB == "YES")
            {
                A_C.FaceContraints(oAssyDoc, oC1, oC2B, 5, 8);
                A_C.AxisContraints(oAssyDoc, oC1, oC2B, 1, 1);
            }
            A_C.AxisContraints(oAssyDoc, oC1, oC3, 1, 1);
            A_C.AxisContraints(oAssyDoc, oC1, oC3, 2, 2);
            A_C.AxisContraints(oAssyDoc, oC1, oC3, 3, 3);

            A_C.AxisContraints(oAssyDoc, oC1, oC4, 1, 1);
            A_C.AxisContraints(oAssyDoc, oC1, oC4, 2, 2);
            A_C.AxisContraints(oAssyDoc, oC1, oC4, 3, 3);



            oAssyDoc.SaveAs("C:\\Rahul\\Nozzle\\NozzleAssemblyflat.iam", false);
        }

    }
}
