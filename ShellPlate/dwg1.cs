using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace ShellPlate
{
    class dwg1
    {

        public void dWG_1(Inventor.Application ThisApplication,string f,double Radius,int level,double N,double[,] coord,double[] H,double[] Length,string name,double[] Thickness,double[] sheetsize,double[] Location,double scl,string[] material,string[] note,string[] sdiscription ,double alfa)
        {
            DrawingDocument oDrawingDoc;
            oDrawingDoc = (DrawingDocument)ThisApplication.ActiveDocument;
            PartDocument oPartDoc;
            oPartDoc = (PartDocument)ThisApplication.Documents.Open(name, false);
            //'Set oDrawingDoc = ThisApplication.Documents.Open(f + "Dwg1.dwg", True)
            // 'Set oDrawingDoc = ThisApplication.Documents.Add(kDrawingDocumentObject,A3
            //Debug.Print Sheet
            Sheet oSheet;
            oSheet = oDrawingDoc.Sheets[1];
            DrawingSheetSizeEnum size;
            //'size = oSheet.size
            //'oSheet.size = kA4DrawingSheetSize
            //'Dim scl As Double
            function obj = new function();
            Dtable objT = new Dtable();
            double offset;
            Double lenght;
            Double width;
            Double SheetLength;
            Double sheetHeight;
            oSheet.Width = 13.38 * 2.54;
            oSheet.Height = 8.1875 * 2.54;
            SheetLength = sheetsize[1];
            sheetHeight = sheetsize[2];

            //'Scl = 0.1
            //'Length = 56.07 * 2.54
            //'width = 23.62 * 2.54
            Point2d oPoint1;
            oPoint1 = ThisApplication.TransientGeometry.CreatePoint2d(Location[1], Location[2]);

            DrawingView oView1;
            DrawingViewLabel oLabel;
                    if(name == (f + "shell.ipt")) 
                    {
        
                    oView1 = oSheet.DrawingViews.AddBaseView((_Document)oPartDoc, oPoint1, scl, ViewOrientationTypeEnum.kBackViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);
                    oView1.ShowLabel = true;
                    oView1.Name = "SHELL ROLLOUT";
                    //'oLabel = "<StyleOverride underline='true'>" & oView1.name & "</styleOverride"
                    oView1.Label.FormattedText = oView1.Name;
                //Creation of table is pending
                    objT.CRtable1(ThisApplication, level, N, coord, H, Length, name, Thickness, sdiscription, material, note, Radius, alfa);
                //Creation of table is pending
            }
                  else
                  {
                    oView1 = oSheet.DrawingViews.AddBaseView((_Document)oPartDoc, oPoint1, scl, ViewOrientationTypeEnum.kTopViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);
                    oView1.ShowLabel = true;
                     oView1.Name = "CHORD DIMENSION AT BOTTOM RING AS NOTED";
                    oView1.Label.FormattedText = oView1.Name;
                  }
                   
        }

        public void dWG_2(Inventor.Application ThisApplication, string f, double Radius, int level, double N, double[,] coord, double[] H, double[] Length, string names, double[] Thickness, double[] sheetsize, double[] Location, double scl, string[] material, string[] note, string[] sdiscription, double alfa)
        {
            DrawingDocument oDrawingDoc;
            oDrawingDoc = (DrawingDocument)ThisApplication.ActiveDocument;
            PartDocument oPartDoc;
            oPartDoc = (PartDocument)ThisApplication.Documents.Open(names, false);
            //'Set oDrawingDoc = ThisApplication.Documents.Open(f + "Dwg1.dwg", True)
            // 'Set oDrawingDoc = ThisApplication.Documents.Add(kDrawingDocumentObject,A3
            //Debug.Print Sheet
            Sheet oSheet;
            oSheet = oDrawingDoc.Sheets[2];
            DrawingSheetSizeEnum size;
            function obj = new function();
            Dtable objT = new Dtable(); 
            //'size = oSheet.size
            //'oSheet.size = kA4DrawingSheetSize
            //'Dim scl As Double
            double offset;
            Double lenght;
            Double width;
            Double SheetLength;
            Double sheetHeight;
            oSheet.Width = 13.38 * 2.54;
            oSheet.Height = 8.1875 * 2.54;
            SheetLength = sheetsize[1];
            sheetHeight = sheetsize[2];
            int[] ar = new int[5];

            //'Scl = 0.1
            //'Length = 56.07 * 2.54
            //'width = 23.62 * 2.54
            Point2d oPoint1;
            oPoint1 = ThisApplication.TransientGeometry.CreatePoint2d(Location[1], Location[2]);

            DrawingView oView1;
            DrawingViewLabel oLabel;
            if (names == (f + "shellC.ipt"))
            {

                oView1 = oSheet.DrawingViews.AddBaseView((_Document)oPartDoc, oPoint1, scl, ViewOrientationTypeEnum.kBackViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);
                oView1.ShowLabel = true;
                obj.toftin(Radius / 30.48, ref ar);

                if(((double)ar[3] / (double)ar[4]) <= 0.0625) {
                    oView1.Name = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + "''" + " INSIDE ROLL RADIUS"; 
                }
                else {
                    oView1.Name = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + " " + Convert.ToString(ar[3]) + "/" + Convert.ToString(ar[4]) + "''" + " INSIDE ROLL RADIUS";
                }
                //'oLabel = "<StyleOverride underline='true'>" & oView1.name & "</styleOverride"
                oView1.Label.FormattedText = oView1.Name;
                //Creation of table is pending
                //CRtable1(level, N, coord, H, Length, name, Thickness, sdiscription, material, note, Radius, alfa);
                //Creation of table is pending
                objT.CRtable2(ThisApplication,level, N, coord, H, Length, names, Thickness, sdiscription, material, note, Radius, alfa);
            }
            else
            {
                oView1 = oSheet.DrawingViews.AddBaseView((_Document)oPartDoc, oPoint1, scl, ViewOrientationTypeEnum.kTopViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);
                //oView1.ShowLabel = true;
                //oView1.Name = "CHORD DIMENSION AT BOTTOM RING AS NOTED";
                //oView1.Label.FormattedText = oView1.Name;
                objT.CRtable3(ThisApplication, level, N, coord, H, Length, names, Thickness, sdiscription, material, note, Radius, alfa);
            }

        }




    }
}
