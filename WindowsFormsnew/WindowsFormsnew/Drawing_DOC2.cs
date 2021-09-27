using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace WindowsFormsnew
{
    class Drawing_DOC2
    {
        public Drawing_DOC2(Inventor.Application InventorApplication, double length, double height, string[] Ship_Discription, string[] Discription, string[] MATL, double[] LGTH, double[] weight, double[] qty, double RollRadius, String[] MK)
        {
            functions DWGF = new functions();
            Double[] sheetsize = new Double[3];
            //osheet.Width = 13.38 * 2.54;
            sheetsize[1] = 13.38 * 2.54;
            sheetsize[2] = 8.1875 * 2.54;
            //osheet.Height = 8.1875 * 2.54;
            DWGF.GetdrawingDoc(InventorApplication, sheetsize);
            DrawingDocument oDrawingDoc;
            //oDrawingDoc = (DrawingDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, InventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject), false);
            oDrawingDoc = (DrawingDocument)InventorApplication.ActiveDocument;
            Sheet osheet, osheet2;
            osheet = oDrawingDoc.Sheets[1];
            osheet2 = oDrawingDoc.Sheets[2];
            double scl;
            scl = DWGF.scls(sheetsize, length, height);


            Point2d oPoint1;
            oPoint1 = InventorApplication.TransientGeometry.CreatePoint2d(7 * 2.54, sheetsize[2] / 2);

            AssemblyDocument OAssmDoc;

            OAssmDoc = (AssemblyDocument)InventorApplication.Documents.Open("C:\\Rahul\\Nozzle\\NozzleAssembly.iam", false);


            DrawingView oView1;
            oView1 = (DrawingView)osheet.DrawingViews.AddBaseView((_Document)OAssmDoc, oPoint1, scl, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle);
            oView1.ShowLabel = true;

            double c;
            functions UC = new functions();
            int[] ar = new int[4];

            UC.toftin(RollRadius, ref ar);
            c = ((double)ar[2]) / ((double)ar[3]);
            string trialString;
            if (c <= 0.0625)
            {
                trialString = Convert.ToString(ar[0]) + "'-" + Convert.ToString(ar[1]) + "''";
            }
            else
            {
                trialString = Convert.ToString(ar[0]) + "'-" + Convert.ToString(ar[1]) + " " + Convert.ToString(ar[2]) + "/" + Convert.ToString(ar[3]) + "''";
            }


            oView1.Name = "Roll Radius " + trialString;
            oView1.Label.FormattedText = oView1.Name;
            OAssmDoc.Close((bool)true);

            Point2d oPoint2, oPoint3, oPoint4;
            oPoint2 = InventorApplication.TransientGeometry.CreatePoint2d(0, height / 2 + 1);
            oPoint3 = InventorApplication.TransientGeometry.CreatePoint2d(0, -height / 2 - 1.5);
            oPoint4 = InventorApplication.TransientGeometry.CreatePoint2d(3 * 2.54, 0);

            DrawingSketch oDrawingSketch;
            oDrawingSketch = oView1.Sketches.Add();
            oDrawingSketch.Edit();
            SketchLine oSketchLine;
            oSketchLine = oDrawingSketch.SketchLines.AddByTwoPoints(oPoint2, oPoint3);
            oDrawingSketch.ExitEdit();

            SectionDrawingView oView2;
            oView2 = osheet.DrawingViews.AddSectionView(oView1, oDrawingSketch, oPoint4, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle);
            oView2.Name = "Elevation";
            oView2.Label.FormattedText = oView2.Name;
            CTable(InventorApplication, sheetsize, Ship_Discription, Discription, MATL, LGTH, weight, qty, MK);
            CTable2(InventorApplication, sheetsize, Ship_Discription, Discription, MATL, LGTH, weight, qty);
            // oView2 = osheet2.DrawingViews.AddSectionView(oView1, oDrawingSketch, oPoint4, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle);
            //Section View
            SheetText(InventorApplication, sheetsize);
            //AssemblyDocument oAssyDoc;
            //oAssyDoc = (AssemblyDocument)InventorApplication.Documents.Add("C:\\Rahul\\Nozzle\\C3.ipt",,false);
            TransientGeometry oTransGeom = default(TransientGeometry);

            oTransGeom = InventorApplication.TransientGeometry;

            // Matrix oPositionMatrix, oPositionMatrix2;

            //oPositionMatrix = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
            //string sfilename = "C:\\Rahul\\Nozzle\\C3.ipt";
            ComponentOccurrence oC;
            //oC3 = oAssmDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);
            if (qty[3] <= 1)
            {
                oC = OAssmDoc.ComponentDefinition.Occurrences[1];
                oView1.SetVisibility(oC, false);
                oC = OAssmDoc.ComponentDefinition.Occurrences[2];
                oView1.SetVisibility(oC, false);
                oC = OAssmDoc.ComponentDefinition.Occurrences[4];
                oView1.SetVisibility(oC, false);
            }
            else
            {
                oC = OAssmDoc.ComponentDefinition.Occurrences[1];
                oView1.SetVisibility(oC, false);
                oC = OAssmDoc.ComponentDefinition.Occurrences[2];
                oView1.SetVisibility(oC, false);
                oC = OAssmDoc.ComponentDefinition.Occurrences[3];
                oView1.SetVisibility(oC, false);
                oC = OAssmDoc.ComponentDefinition.Occurrences[5];
                oView1.SetVisibility(oC, false); 
            }



            SheetText2(InventorApplication, sheetsize);
            oDrawingDoc.Save();


        }


        public void CTable(Inventor.Application InventorApplication, Double[] sheetsize, string[] Ship_Discription, string[] Discription, string[] MATL, double[] LGTH, double[] weight, double[] qty, string[] MK)
        {
            Double length;
            functions UC = new functions();
            int[] ar = new int[4];

            DrawingDocument oDrawDoc;
            oDrawDoc = (DrawingDocument)InventorApplication.ActiveDocument;

            Sheet osheet;
            osheet = oDrawDoc.Sheets[1];

            int row = 9;
            string[] oTitles = new string[row];
            oTitles[0] = "SA/PC";
            oTitles[1] = "MK";
            oTitles[2] = "QTY";
            oTitles[3] = "SHIP DESCRIPTION";
            oTitles[4] = "DESCRIPTION";
            oTitles[5] = "LGTH";
            oTitles[6] = "MATL";
            oTitles[7] = "Notes";
            oTitles[8] = "WT/EA";

            string[] oContents;
            int flag;
            if (qty[3] > 1)
            {
                oContents = new string[row * 3];
                flag = 3;
            }
            else
            {
                oContents = new string[row * 2];
                flag = 2;
            }





            int k = 0;
            for (int i = 0; i < row * (qty[3] + 1); i++)
            {
                if ((i % row) == 0)
                {
                    k = k + 1;
                    oContents[i] = Convert.ToString(k);
                }

                if ((i % row) == 1)
                {

                    oContents[i] = MK[k - 1];
                }

                if ((i % row) == 2)
                {
                    oContents[i] = Convert.ToString(1);
                }

                if ((i % row) == 3)
                {
                    //sdiscription
                    oContents[i] = Ship_Discription[k - 1];
                }

                if ((i % row) == 4)
                {
                    //sdiscription
                    oContents[i] = Discription[k - 1];

                }

                if ((i % row) == 5)
                {
                    double c;
                    UC.toftin(LGTH[k - 1], ref ar);
                    c = ((double)ar[2]) / ((double)ar[3]);
                    if (c <= 0.0625)
                    {

                        oContents[i] = Convert.ToString(ar[0]) + "'-" + Convert.ToString(ar[1]) + "''";
                    }
                    else
                    {
                        oContents[i] = Convert.ToString(ar[0]) + "'-" + Convert.ToString(ar[1]) + " " + Convert.ToString(ar[2]) + "/" + Convert.ToString(ar[3]) + "''";
                    }

                    Console.WriteLine("ar 0 {0},1 {1},2 {2},3 {3} ", ar[0], ar[1], ar[2], ar[3]);
                    //sdiscription


                }

                if ((i % row) == 6)
                {
                    //sdiscription
                    oContents[i] = MATL[k - 1];

                }

                if ((i % row) == 7)
                {
                    //sdiscription
                    oContents[i] = "Notes";

                }

                if ((i % row) == 8)
                {
                    //sdiscription
                    oContents[i] = Convert.ToString(weight[k - 1]);


                }


            }

            Double[] oColumnWidths = new double[row];
            oColumnWidths[0] = 2.54 * 0.4;
            oColumnWidths[1] = 2.54 * 0.4;
            oColumnWidths[2] = 2.54 * 0.4;
            oColumnWidths[3] = 3 * 2.54 * 0.4;
            oColumnWidths[4] = 4 * 2.54 * 0.4;
            oColumnWidths[5] = 3 * 2.54 * 0.4;
            oColumnWidths[6] = 3 * 2.54 * 0.4;
            oColumnWidths[7] = 2 * 2.54 * 0.4;
            oColumnWidths[8] = 2 * 2.54 * 0.4;

            CustomTable oCustomTable;
            oCustomTable = osheet.CustomTables.Add("BILL OF MATERIAL(US UNITS)", InventorApplication.TransientGeometry.CreatePoint2d(6 * 2.54, 10 * 2.54), row, flag, oTitles, oContents, oColumnWidths);

            Double Oheight;
            Oheight = (oCustomTable.RangeBox.MaxPoint.Y - oCustomTable.RangeBox.MinPoint.Y) / 2.54;

            Double oLength;
            oLength = (oCustomTable.RangeBox.MaxPoint.X - oCustomTable.RangeBox.MinPoint.X) / 2.54;

            oCustomTable.Position = InventorApplication.TransientGeometry.CreatePoint2d(sheetsize[1] - oLength * 2.54 - 0.394 * 2.54, sheetsize[2] - 0.394 * 2.54);
            oDrawDoc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.TableStyle.RowGap = 0.1;

            TableFormat oFormat;
            oFormat = osheet.CustomTables.CreateTableFormat();
            TextStyle oTextstyles;
            oTextstyles = oDrawDoc.StylesManager.TextStyles[18];
            oTextstyles.Font = "Arial Black";
            oTextstyles.FontSize = 0.15;
            oFormat.TextStyle = oTextstyles;
            oFormat.OutsideLineWeight = 0.05;
            oFormat.InsideLineWeight = 0.05;
            oCustomTable.OverrideFormat = oFormat;


        }


        public void CTable2(Inventor.Application InventorApplication, Double[] sheetsize, string[] Ship_Discription, string[] Discription, string[] MATL, double[] LGTH, double[] weight, double[] qty)
        {
            Double length;
            functions UC = new functions();
            int[] ar = new int[4];

            DrawingDocument oDrawDoc;
            oDrawDoc = (DrawingDocument)InventorApplication.ActiveDocument;

            Sheet osheet;
            osheet = oDrawDoc.Sheets[2];

            int row = 7;
            string[] oTitles = new string[row];
            oTitles[0] = "SA/PC";
            //oTitles[1] = "MK";
            oTitles[1] = "QTY";
            //oTitles[2] = "SHIP DESCRIPTION";
            oTitles[2] = "DESCRIPTION";
            oTitles[3] = "LGTH";
            oTitles[4] = "MATL";
            oTitles[5] = "Notes";
            oTitles[6] = "WT/EA";

            string[] oContents = new string[row * 2];
            int k = 3;
            for (int i = 0; i < row * 2; i++)
            {
                if ((i % row) == 0)
                {
                    k = k + 1;
                    oContents[i] = Convert.ToString(k);
                }

                //if ((i % row) == 1)
                //{

                //    oContents[i] = "SR"+ Convert.ToString(k);
                //}

                if ((i % row) == 1)
                {
                    oContents[i] = Convert.ToString(1);
                }

                //if ((i % row) == 2)
                //{
                //    sdiscription
                //    oContents[i] = Ship_Discription[k - 1];
                //}

                if ((i % row) == 2)
                {
                    //sdiscription
                    oContents[i] = Discription[k - 1];

                }

                if ((i % row) == 3)
                {
                    double c;


                    UC.toftin(LGTH[k - 1], ref ar);
                    c = ((double)ar[2]) / ((double)ar[3]);
                    if (c <= 0.0625)
                    {
                        oContents[i] = Convert.ToString(ar[0]) + "'-" + Convert.ToString(ar[1]) + "''";
                    }
                    else
                    {
                        oContents[i] = Convert.ToString(ar[0]) + "'-" + Convert.ToString(ar[1]) + " " + Convert.ToString(ar[2]) + "/" + Convert.ToString(ar[3]) + "''";
                    }
                    //sdiscription
                    Console.WriteLine("ar 0 :{0},1: {1},2 :{2},3 :{3},c: {4} ", ar[0], ar[1], ar[2], ar[3], c);

                }

                if ((i % row) == 4)
                {
                    //sdiscription
                    oContents[i] = MATL[k - 1];

                }

                if ((i % row) == 5)
                {
                    //sdiscription
                    oContents[i] = "Notes";

                }

                if ((i % row) == 6)
                {
                    //sdiscription
                    oContents[i] = Convert.ToString(weight[k - 1]);


                }


            }

            Double[] oColumnWidths = new double[row];
            oColumnWidths[0] = 2.54 * 0.4;
            //oColumnWidths[1] = 2.54 * 0.4;
            oColumnWidths[1] = 2.54 * 0.4;
            oColumnWidths[2] = 4 * 2.54 * 0.4;
            //   oColumnWidths[3] = 4 * 2.54 * 0.4;
            oColumnWidths[3] = 3 * 2.54 * 0.4;
            oColumnWidths[4] = 3 * 2.54 * 0.4;
            oColumnWidths[5] = 2 * 2.54 * 0.4;
            oColumnWidths[6] = 2 * 2.54 * 0.4;

            CustomTable oCustomTable;
            oCustomTable = osheet.CustomTables.Add("BILL OF MATERIAL(US UNITS)", InventorApplication.TransientGeometry.CreatePoint2d(6 * 2.54, 10 * 2.54), row, 2, oTitles, oContents, oColumnWidths);

            Double Oheight;
            Oheight = (oCustomTable.RangeBox.MaxPoint.Y - oCustomTable.RangeBox.MinPoint.Y) / 2.54;

            Double oLength;
            oLength = (oCustomTable.RangeBox.MaxPoint.X - oCustomTable.RangeBox.MinPoint.X) / 2.54;

            oCustomTable.Position = InventorApplication.TransientGeometry.CreatePoint2d(sheetsize[1] - oLength * 2.54 - 0.394 * 2.54, sheetsize[2] - 0.394 * 2.54);
            oDrawDoc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.TableStyle.RowGap = 0.1;

            TableFormat oFormat;
            oFormat = osheet.CustomTables.CreateTableFormat();
            TextStyle oTextstyles;
            oTextstyles = oDrawDoc.StylesManager.TextStyles[18];
            oTextstyles.Font = "Arial Black";
            oTextstyles.FontSize = 0.15;
            oFormat.TextStyle = oTextstyles;
            oFormat.OutsideLineWeight = 0.05;
            oFormat.InsideLineWeight = 0.05;
            oCustomTable.OverrideFormat = oFormat;

        }

        public void SheetText(Inventor.Application InventorApplication, Double[] sheetsize)
        {
            DrawingDocument oDrawdoc;
            oDrawdoc = (DrawingDocument)InventorApplication.ActiveDocument;
            Sheet oActiveSheet;
            oActiveSheet = oDrawdoc.Sheets[1];
            GeneralNotes oGeneralNotes;
            oGeneralNotes = oActiveSheet.DrawingNotes.GeneralNotes;
            TransientGeometry oTG;
            oTG = InventorApplication.TransientGeometry;

            String sText;
            Double dYcoord, CN;
            CN = sheetsize[2] / 2.54 - 2;
            Double x1;
            x1 = sheetsize[1] / (2.54) - 3.18;
            Double x2;
            x2 = x1 + 0.2;
            dYcoord = CN * 2.54;
            sText = "NOTE";
            GeneralNote oGeneralNote;
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, CN * 2.54), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + 0.5);

            Double dYoffset;
            TextStyle oStyle;
            oStyle = oGeneralNotes[1].TextStyle;
            oStyle.FontSize = 0.25;
            dYoffset = oStyle.FontSize * 1.1;
            Double gap;
            gap = 0.2;
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "1.");
            sText = "WPG-3 UNLESS NOTED";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "2.");
            sText = "NDE-LT1(TWICE) UNLESS NOTED";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "3.");
            sText = "TELL TALE HOLE TO BE PACKED WITH";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            sText = "HEAVY GREASE AFTER TESTING";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);



            //Drawing

        }

        public void SheetText2(Inventor.Application InventorApplication, Double[] sheetsize)
        {
            DrawingDocument oDrawdoc;
            oDrawdoc = (DrawingDocument)InventorApplication.ActiveDocument;
            Sheet oActiveSheet;
            oActiveSheet = oDrawdoc.Sheets[2];
            GeneralNotes oGeneralNotes;
            oGeneralNotes = oActiveSheet.DrawingNotes.GeneralNotes;
            TransientGeometry oTG;
            oTG = InventorApplication.TransientGeometry;

            String sText;
            Double dYcoord, CN;
            CN = sheetsize[2] / 2.54 - 2;
            Double x1;
            x1 = sheetsize[1] / (2.54) - 3.18;
            Double x2;
            x2 = x1 + 0.2;
            dYcoord = CN * 2.54;
            sText = "NOTE";
            GeneralNote oGeneralNote;
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, CN * 2.54), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + 0.5);

            Double dYoffset;
            TextStyle oStyle;
            oStyle = oGeneralNotes[1].TextStyle;
            oStyle.FontSize = 0.25;
            dYoffset = oStyle.FontSize * 1.1;
            Double gap;
            gap = 0.2;
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "1.");
            sText = "WPG-2";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "2.");
            sText = "NDE-VT1";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);



            //Drawing

        }
    }
}
