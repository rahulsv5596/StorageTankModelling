using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.IO;

namespace ShellPlate
{
    class shell_2d
    {
        public void Shell2d(Inventor.Application InventorApplication, int level1, double Radius, double N1, double[] H, double[] Thickness, double dt, string[] material, string[] note, string[] sdiscription)
        {
            Double R1;
            Double R2;
            function obj = new function();
            int level = level1;
            int N = (int)N1;

            TransientGeometry oTransGeom;
            oTransGeom = InventorApplication.TransientGeometry;

            PartDocument oPartDoc;


            oPartDoc = (PartDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject), true);

            PlanarSketch oSketch;

            Double Pi;
            Pi = Math.PI;
            PartComponentDefinition oPartCompDef;
            oPartCompDef = (PartComponentDefinition)oPartDoc.ComponentDefinition;


            oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes[2]);


            Profile oProfile;
            ExtrudeFeature oExtrude;

            int i = 0;
            Double angleshift;
            angleshift = (2 * Pi) / (3 * N);
            R2 = Radius;
            Double[] Length = new double[level+1];
            Double[] Length2 = new double[level+1];
            Double delta;
            delta = dt / N;

            //Dim Length2() As Double
            //ReDim Length2(level)
            //'length = 2 * Pi * Radius / N

            Double xcoord;
            Double[,] coord = new double[N + 2, level+1];
            //ReDim coord(N +1, level)
            xcoord = 0;

            //loop for making 2-D shell;
            for (int j = 1; j <= level; j++) {
                R1 = R2 + Thickness[j] / 2;
                Length[j] = (2 * Pi * (R1 + Thickness[j] / 2)) / N - dt / N;
                Length2[j] = (2 * Pi * R1) / N - dt / N;

                    for(i = 1; i <= N; i++) {
                           // oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], H[j - 1]));
                             oSketch = (PlanarSketch)oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes[2], H[j - 1]));
                             Point2d oCoord1;
                             oCoord1 = oTransGeom.CreatePoint2d(xcoord - dt / (2 * N), 0);
                        //'Set oCoord1 = oTransGeom.CreatePoint2d(R1 * Cos(((((i - 1) * 2 * Pi) + dt) / N + (j - 1) * angleshift)), R1 * Sin(((((i - 1) * 2 * Pi + dt)) / N + (j - 1) * angleshift)))
                             Point2d oCoord2;
                             oCoord2 = oTransGeom.CreatePoint2d(xcoord + Length[j] + dt / (2 * N), Thickness[j]);
                             oSketch.SketchLines.AddAsTwoPointRectangle(oCoord1, oCoord2);
                             oProfile = oSketch.Profiles.AddForSolid();

                      
                    if (level == 1)
                    {



                        //oExtrude = oPartCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, 0.1, kPositiveExtentDirection, kNewBodyOperation);
                        function oExtd = new function();
                        oExtd.extrude(oPartCompDef, oProfile, 0.1, 2, 2);
                    }
                    else
                    {
                        function oExtd = new function();
                        oExtd.extrude(oPartCompDef, oProfile, H[j] - H[j - 1] - Math.Round(delta, 3), 2, 2);
                        //oExtrude = oPartCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, H[j] - H[j - 1] - Math.Round(delta, 3), kPositiveExtentDirection, kNewBodyOperation);
                    }

                    coord[i, j] = xcoord;
                    xcoord = xcoord + Length[j] + dt / N;
                    }
                    if(i > N) {
                        coord[i, j] = xcoord;
;                    }
                xcoord = (j % 3) * Length[j] / 3;
            }

            string name;
            double scl;
            string f;
            f = "C:\\Rahul\\shell\\";
            if (Directory.Exists(f) == false)
            {

                Directory.CreateDirectory(f);

            }

            double alfa;
            oPartDoc.SaveAs(f + "shell.ipt", false);
            //double alfa;
            obj.MakeComponentsProgrammatically(InventorApplication,f, f + "shell.ipt", f + "shellC.ipt", 1, 1);
            CBA Newobj = new CBA();
            Dtable objT = new Dtable();
            dwg1 objD = new dwg1();
            //oPartDoc = (PartDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject), true);
            alfa = Newobj.CircularBoundaryArray(InventorApplication, level, Radius, N, H, Thickness, dt, f);
            //oPartDoc.SaveAs(f + "CBA.ipt", false);
            name = oPartDoc.FullFileName;
            obj.MakeComponentsProgrammatically(InventorApplication, f, f + "CBA.ipt", f + "CBAC.ipt", 1, 1);
            Double[] sheetsize = new double[3];
            sheetsize[1] = 13.38 * 2.54;
            sheetsize[2] = 8.1875 * 2.54;

            Double[] Location = new double[3];
            Location[1] = sheetsize[1] / 2;
            Location[2] = sheetsize[2] / 3;
            name = f + "shell.ipt";
            //
            scl = obj.scls(Radius, level, N, coord, H, Length2, name, Thickness, sheetsize, 2);
            //
            obj.GetdrawingDoc(InventorApplication,f,2,sheetsize);
            objD.dWG_1(InventorApplication,f, Radius, level, N, coord, H, Length2, name, Thickness, sheetsize, Location, scl, material, note, sdiscription, alfa);
            Location[1] = sheetsize[1] / 2;
            Location[2] = 2 * sheetsize[2] / 3;
            //'scl = 1
            name = f + "CBAC.ipt";
            scl = obj.scls(Radius, level, N, coord, H, Length2, name, Thickness, sheetsize, 1);

            objD.dWG_1(InventorApplication, f, Radius, level, N, coord, H, Length2, name, Thickness, sheetsize, Location, scl, material, note, sdiscription, alfa);
            Location[1] = 10 * 2.54;
            Location[2] = 3 * 2.54;
            name = f + "CBAC.ipt";
            scl = scl / 2;

            objD.dWG_2(InventorApplication, f, Radius, level, N, coord, H, Length2, name, Thickness, sheetsize, Location, scl, material, note, sdiscription, alfa);
            Location[1] = 3.5 * 2.54;
            Location[2] = 5 * 2.54;
            name = f + "shellC.ipt";

            objD.dWG_2(InventorApplication, f, Radius, level, N, coord, H, Length2, name, Thickness, sheetsize, Location, scl, material, note, sdiscription, alfa);
            objT.CRtable4(InventorApplication, level, N, coord, H, Length2, name, Thickness, sdiscription, material, note, Radius, alfa);

            sheetext objS = new sheetext();
            objS.SheetTextAdd1(InventorApplication, f, sheetsize);


        }

    }
}
