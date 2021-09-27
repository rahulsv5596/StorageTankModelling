using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.IO;

namespace ShellPlate
{
    class function
    {


        public void GetdrawingDoc(Inventor.Application InventorApplication,string f,int Tsheet, Double[] Sheetsize)
        {
            DrawingDocument oDrawingDoc;
            // oDrawingDoc = (DrawingDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject), false);
            oDrawingDoc = (DrawingDocument)InventorApplication.Documents.Open("C:\\Rahul\\matrixdwg.dwg", true);
            oDrawingDoc.SaveAs(f+"Shelldwg.dwg", false);
            Sheet osheet;
            String[] sPromptStrings = new String[2];
            sPromptStrings[0] = "String 1";
            sPromptStrings[1] = "String 2";
            TitleBlockDefinition oTitleBlockDef;
            oTitleBlockDef = oDrawingDoc.TitleBlockDefinitions[3];
            TitleBlock oTitleBlock;
            //oTitleBlock.Location = TitleBlockLocationEnum.kBottomLeftPosition;
            for (int i = 1; i <= Tsheet; i++)
            {
            osheet = (Sheet)oDrawingDoc.Sheets.Add(DrawingSheetSizeEnum.kCustomDrawingSheetSize, PageOrientationTypeEnum.kLandscapePageOrientation, "Sheet", Sheetsize[1], Sheetsize[2]);
            osheet.AddTitleBlock(oTitleBlockDef);
            osheet.AddDefaultBorder(8, BorderLabelModeEnum.kBorderLabelModeNumeric, 4, BorderLabelModeEnum.kBorderLabelModeAlphabetical);
            }


        }

        public double scls(double Radius,int level,double N,double[,] coord,double[] H, double[] Length, String name,double[] Thickness,double[] sheetsize,int count)
        {
            Double max=0, min=0, Olen, ohgt;
            //Olen = Length;
            //ohgt = Height;
            double scl;
            if (count > 1)
            {
                max = 0;
                min = 10000000;

                for (int j = 1; j <= level; j++)
                {
                    for (int i = 1; i <= (N + 1); i++)
                    {

                        if (coord[i, j] > max)
                        {
                            max = coord[i, j];
                        }

                        if (coord[i, j] < min)
                        {
                            min = coord[i, j];
                        }


                    }

                }
                Olen = max - min;
                ohgt = H[level];
            }
            else
            {
                Olen = Length[1];
                ohgt = H[1];

            }

            if (Olen > sheetsize[1])
            {
                scl = 0.6 * sheetsize[1] / Olen;

            }
            else
            {
                scl = 0.8;
            }

            if ((scl * ohgt) > sheetsize[2])
            {
                scl = 0.6 * sheetsize[2] / ohgt;

            }

            return scl;

        }

        public void toftin(double value, ref int[] ar)
        {
            //value = value / 2.54;
            //value = value / 12;
            double a, b, c, d, e, X;
            a = (int)value;
            b = value - a;
            b = b * 12;
            c = (int)b;
            d = b - c;
            //d = (int)(d * 16);
            double temp;
            temp = (int)(d * 16);
            Console.WriteLine("{0}:{1},", value, (d * 16));
            temp = Math.Abs(d * 16 - temp);

            d = (int)(d * 16);
            if (temp > 0.2)
            {
                d = d + 1;
            }
            e = 16; ;
            X = gcd(d, e);
            d = d / X;
            e = e / X;
            ar[1] = (int)a;
            ar[2] = (int)c;
            ar[3] = (int)d;
            ar[4] = (int)e;

        }

        public Double VolumeC(Inventor.Application InventorApplication, string f, int i)
        {
            PartDocument partdoc;
            partdoc = (PartDocument)InventorApplication.Documents.Open(f);
            Double VC;
            SurfaceBody oBody;
            // PartComponentDefinition partdef;
            oBody = partdoc.ComponentDefinition.SurfaceBodies[i];
            VC = oBody.Volume[0.001];
            return VC;
        }

        public double MAx_Area(Inventor.Application InventorApplication, string f, int i)
        {
            PartDocument partdoc;
            partdoc = (PartDocument)InventorApplication.Documents.Open(f);
            Double AC = 0;
            //Face oface;
            SurfaceBody oBody;

            oBody = partdoc.ComponentDefinition.SurfaceBodies[1];
            SurfaceEvaluator oFaceEvaluator;
            foreach (Face oface in partdoc.ComponentDefinition.SurfaceBodies[1].Faces)
            {
                oFaceEvaluator = oface.Evaluator;
                if (oFaceEvaluator.Area > AC)
                {
                    AC = oFaceEvaluator.Area;
                }

            }

            return AC;

        }

        public int gcd(double a, double b)
        {
            if (a == 0)
            {
                return (int)b;
            }
            else if (b == 0)
            {
                return (int)a;
            }
            else if (a < b)
            {
                return gcd(a, (int)b % (int)a);

            }
            else
            {
                return gcd(b, (int)a % (int)b);

            }
        }
        public int findmax(int num1, int num2)
        {
            int result;
            if (num1 > num2)
                result = num1;
            else
                result = num2;
            return result;
        }

        public void extrude(PartComponentDefinition oPartCompDef, Profile oProfile, double length, int j, int i)
        {
            ExtrudeDefinition oextrudedef = default(ExtrudeDefinition);

            if (j == 1)
            {
                oextrudedef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kJoinOperation);
            }
            else if (j == 2)
            {
                oextrudedef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kNewBodyOperation);
            }
            else
            {
                oextrudedef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kCutOperation);
            }


            if (i == 1)
            {
                oextrudedef.SetDistanceExtent(length, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            }
            else if (i == 2)
            {
                oextrudedef.SetDistanceExtent(length, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);
            }
            else
            {
                oextrudedef.SetDistanceExtent(length, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
            }


            ExtrudeFeature oExtrude = default(ExtrudeFeature);
            oExtrude = oPartCompDef.Features.ExtrudeFeatures.Add(oextrudedef);
        }

        


        //public void ExcelRead(double C_Value, ref double[] arr, string Path, int Esize, int Csize)
        //{
        //    Excel excel = new Excel(Path, 1);

        //    int i, j;
        //    int C_Index = 0;
        //    double CiV;

        //    double[] N_Parameter = new double[10];
        //    j = 1;
        //    for (i = 2; i <= Esize; i++)
        //    {
        //        //piV = Convert.ToDouble((excel.ReadCell((i - 1), 1)));
        //        CiV = Convert.ToDouble((excel.ReadCell(i, j)));
        //        //Console.WriteLine("entering at {0} {1}", CiV,C_Value);
        //        if (Math.Abs(C_Value - CiV) < 10e-6)
        //        {

        //            C_Index = i;
        //        }

        //    }

        //    for (i = 1; i <= Csize; i++)
        //    {
        //        arr[i] = Convert.ToDouble((excel.ReadCell(C_Index, i)));
        //        arr[i] = arr[i] * 2.54;
        //    }




        //}


        public void AxisContraints(AssemblyDocument oAssyDoc, ComponentOccurrence oC1, ComponentOccurrence oC2, int i, int j)
        {

            Inventor.PartComponentDefinition pPartdef1 = null;
            pPartdef1 = oC1.Definition as PartComponentDefinition;

            Inventor.PartComponentDefinition pPartdef2 = null;
            pPartdef2 = oC2.Definition as PartComponentDefinition;

            Inventor.WorkAxis pPartAxis1 = pPartdef1.WorkAxes[i];
            Inventor.WorkAxis pPartAxis2 = pPartdef2.WorkAxes[j];

            object axisProxy1 = null;
            Inventor.WorkAxisProxy pWorkAxisProxy1 = null;
            oC1.CreateGeometryProxy(pPartAxis1, out axisProxy1);
            pWorkAxisProxy1 = axisProxy1 as WorkAxisProxy;

            object axisProxy2 = null;
            Inventor.WorkAxisProxy pWorkAxisProxy2 = null;
            oC2.CreateGeometryProxy(pPartAxis2, out axisProxy2);
            pWorkAxisProxy2 = axisProxy2 as WorkAxisProxy;

            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oAxisDef;
            oAxisDef = oAssyDoc.ComponentDefinition;
            oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(pWorkAxisProxy1, pWorkAxisProxy2, 0.0);
        }




        public void FaceContraints(AssemblyDocument oAssyDoc, ComponentOccurrence oC1, ComponentOccurrence oC2, int i, int j)
        {

            int k = 1;
            int faceC;
            faceC = oC1.SurfaceBodies[1].Faces.Count;
            //Face[] oCA = new Face[faceC + 1];
            Face oface1, oface2;

            oface1 = oC1.SurfaceBodies[1].Faces[i];
            oface2 = oC2.SurfaceBodies[1].Faces[j];

            AssemblyConstraint oConstr;
            AssemblyComponentDefinition oAxisDef;
            oAxisDef = oAssyDoc.ComponentDefinition;
            oConstr = (AssemblyConstraint)oAxisDef.Constraints.AddMateConstraint(oface1, oface2, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference, null, null);
        }

        public void WriteCustomData(string propName, object propValue, Document invDoc)
        {
            PropertySet propSet = invDoc.PropertySets["Inventor User Defined Properties"];
            Inventor.Property prop;
            try
            {
                prop = propSet[propName];
            }
            catch (Exception e)
            {
                prop = propSet.Add("", propName);
            }
            prop.Value = propValue;


        }

        public void ReadCustomData(string propName, double propValue, Document invDoc)
        {
            PropertySet propSet = invDoc.PropertySets["Inventor User Defined Properties"];
            Inventor.Property prop;

            //prop = propSet.ItemByPropId("Part Number");

            prop = propSet[propName];



        }

        public void MakeComponentsProgrammatically(Inventor.Application ThisApplication,string f, string name, string filename, int BSSN, int BSEN)
        {
            //object fso;
            //Inventor.Application ThisApplication = default(Inventor.Application);
            //fso = ThisApplication.FileManager.FileSystemObject;
            if (Directory.Exists(f) == false)
            {

                Directory.CreateDirectory(f);

            }

            //PartDocument oPartDoc;
            //oPartDoc = (PartDocument)InventorApplication.Documents.Open(f + "shell\\CBA.ipt", true);

            PartDocument doc;
            doc = (PartDocument)ThisApplication.Documents.Open(name, false);
            //oDrawingDoc = (DrawingDocument)InventorApplication.Documents.Open("C:\\Rahul\\matrixdwg.dwg", true);
            PartDocument prt;
            prt = (PartDocument)ThisApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject);
            SurfaceBody sb=default(SurfaceBody);
            //object sb=10;
            // sb = new SurfaceBody();
            for (int i = BSSN; i <= BSEN; i++)
            {
                sb = doc.ComponentDefinition.SurfaceBodies[i];
                DerivedPartComponents dpcs;
                dpcs = (DerivedPartComponents)prt.ComponentDefinition.ReferenceComponents.DerivedPartComponents;
                DerivedPartUniformScaleDef dpd;
                dpd = dpcs.CreateUniformScaleDef(doc.FullDocumentName);
                //dpe;
                foreach (DerivedPartEntity dpe in dpd.Solids)
                {
                    //dpe.IncludeEntity = dpe.ReferencedEntity(sb);
                    if(dpe.ReferencedEntity == sb)
                    dpe.IncludeEntity = true;
                    else
                    dpe.IncludeEntity = false;
                }

                dpcs.Add((DerivedPartDefinition)dpd);
            }

            ThisApplication.SilentOperation = true;
            prt.SaveAs(filename, false);
            ThisApplication.SilentOperation = false;

        }



    }
}
