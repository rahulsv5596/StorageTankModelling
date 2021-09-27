using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace ShellPlate
{
    class Dtable
    {
        public void CRtable1(Inventor.Application ThisApplication,int level,double N,double[,] coord,double[] H,double[] Length,string name,double[] thickness,string[] sdiscription,string[] material,string[] note,double Radius,double alfa)
        {
            DrawingDocument oDrawDoc;
            oDrawDoc = (DrawingDocument)ThisApplication.ActiveDocument;
            int[] ar = new int[5];
            Sheet oSheet;
            oSheet = oDrawDoc.Sheets[1];
             int Row;
            Row = 2;
            string[] oTitles= new string[Row];
            oTitles[0] = "MARK NO.";
            oTitles[1] = "CHORD LENGTH";

            string[] oContents = new string[Row];

            int k = 0;
            int j = 1;

            oContents[0] = "SR1";
            function obj = new function();
            obj.toftin(2 * ((Radius) / 30.48) * Math.Sin(alfa / 2),ref ar);



            if (((double)ar[3] / (double)ar[4]) <= 0.0625) {
                oContents[1] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + "''";
            }
            else
            {
                oContents[1] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + " " + Convert.ToString(ar[3]) + "/" + Convert.ToString(ar[4]) + "''";
            }



        CustomTable oCustomTable;
            oCustomTable = oSheet.CustomTables.Add("BOTTOM RING CHORDS", ThisApplication.TransientGeometry.CreatePoint2d(2 * 2.54, 7.7 * 2.54), Row, 1, oTitles, oContents);
            oDrawDoc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.TableStyle.RowGap = 0.1;
         TableFormat oFormat;
            oFormat = oSheet.CustomTables.CreateTableFormat();
           TextStyle oTextstyles;
            oTextstyles = oDrawDoc.StylesManager.TextStyles[18];
            oTextstyles.Font = "RomanS";
            oTextstyles.FontSize = 0.15;
            oFormat.TextStyle = oTextstyles;

            oFormat.OutsideLineWeight = 0.05;
            oCustomTable.OverrideFormat = oFormat;


        }

        public void CRtable2(Inventor.Application ThisApplication, int level, double N, double[,] coord, double[] H, double[] Length, string name, double[] Thickness, string[] sdiscription, string[] material, string[] note, double Radius, double alfa)
        {
            DrawingDocument oDrawDoc;
            oDrawDoc = (DrawingDocument)ThisApplication.ActiveDocument;
            int[] ar = new int[5];
            Sheet oSheet;
            oSheet = oDrawDoc.Sheets[2];
            int Row;
            Row = 9;
            function obj = new function();
            string[] oTitles = new string[Row];
            oTitles[0] = "SA/PC";
            oTitles[1] = "MK";
            oTitles[2] = "QTY";
            oTitles[3] = "SHIP DESCRIPTION";
            oTitles[4] = "DESCRIPTION";
            oTitles[5] = "LGTH";
            oTitles[6] = "MATL";
            oTitles[7] = "Notes";
            oTitles[8] = "WT/EA";

            string[] oContents = new string[Row*level];
            int k = 0;
            int i;


            for ( i = 0; i < Row * level; i++) {

                if((i % Row) == 0){
                k = k + 1;
                oContents[i] = Convert.ToString(k);

                }

                if(i % Row == 1) {
                    oContents[i] = "SR" + k;

                }

                if(i % Row == 2) {
                    oContents[i] = Convert.ToString(N);
                }

                if(i % Row == 3) {
                    oContents[i] = sdiscription[k];
                }

               if(i % Row == 4) {
                    oContents[i] = "PL " + (Thickness[k]) / 2.54 + " in x " + (H[k] - H[k - 1]) / 2.54 + "''";
                }

                 if(i % Row == 5) {
                    obj.toftin(Length[k]/30.48, ref ar);    
                    //toftin Length(k) / 30.48, ar;
                    if (((double)ar[3] / (double)ar[4]) <= 0.0625) {
                        oContents[i] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + "''";
                    }
                    else {
                        oContents[i] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + " " + Convert.ToString(ar[3]) + "/" + Convert.ToString(ar[4]) + "''";
                    }
                }

                if(i % Row == 6) {
                    oContents[i] = material[k];
                }

                 if(i % Row == 7) {
                    oContents[i] = note[k];
                }

                 if(i % Row == 8) {
                    oContents[i] = Convert.ToString(Math.Round((Length[k] / 30.48) * (Thickness[k] / 2.54) * ((H[k] - H[k - 1]) / 2.54) * 0.2836 * 12, 2));
                }
             }

            Double[] oColumnWidths = new double[Row];
            //ReDim oColumnWidths(1 To Row)
            oColumnWidths[0] = 2.54 * 0.4;
            oColumnWidths[1] = 2.54 * 0.4;
            oColumnWidths[2] = 2.54 * 0.4;
            oColumnWidths[3] = 3 * 2.54 * 0.4;
            oColumnWidths[4] = 3 * 2.54 * 0.4;
            oColumnWidths[5] = 3 * 2.54 * 0.4;
            oColumnWidths[6] = 3 * 2.54 * 0.4;
            oColumnWidths[7] = 2 * 2.54 * 0.4;
            oColumnWidths[8] = 2 * 2.54 * 0.4;

            CustomTable oCustomTable;
            oCustomTable = oSheet.CustomTables.Add("BILL OF MATERIAL (US UNITS)", ThisApplication.TransientGeometry.CreatePoint2d(6 * 2.54, 10 * 2.54), Row, level, oTitles, oContents, oColumnWidths);

        

            oCustomTable.Columns[3].ValueHorizontalJustification= HorizontalTextAlignmentEnum.kAlignTextCenter;


            double[] sheetsize=new double[3];
            sheetsize[1] = 13.38 * 2.54;
            sheetsize[2] = 8.1875 * 2.54;


            double Oheight;
            Oheight = (oCustomTable.RangeBox.MaxPoint.Y - oCustomTable.RangeBox.MinPoint.Y) / 2.54;



             double oLength;
            oLength = (oCustomTable.RangeBox.MaxPoint.X - oCustomTable.RangeBox.MinPoint.X) / 2.54;
            oCustomTable.Position = ThisApplication.TransientGeometry.CreatePoint2d(sheetsize[1] - oLength * 2.54 - 0.394 * 2.54, sheetsize[2] - 0.394 * 2.54);
           // oDrawDoc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.TableStyle.RowGap = 0.1;


            TableFormat oFormat;
            oFormat = oSheet.CustomTables.CreateTableFormat();
             TextStyle oTextstyles;
            oTextstyles = oDrawDoc.StylesManager.TextStyles[18];
            oTextstyles.Font = "Arial Black";
            oTextstyles.FontSize = 0.15;
            oFormat.TextStyle = oTextstyles;
            oFormat.OutsideLineWeight = 0.05;
            oCustomTable.OverrideFormat = oFormat;

        }

        public void CRtable3(Inventor.Application ThisApplication, int level, double N, double[,] coord, double[] H, double[] Length, string name, double[] Thickness, string[] sdiscription, string[] material, string[] note, double Radius, double alfa)
        {
            DrawingDocument oDrawDoc;
            oDrawDoc = (DrawingDocument)ThisApplication.ActiveDocument;
            int[] ar = new int[5];
            Sheet oSheet;
            oSheet = oDrawDoc.Sheets[2];
            int Row;
            Row = 5;
            Double[] dimention1 = new double[level+1];
            Double[] dimention2 = new double[level+1];
  
            Double delta;
             Double dt;
            dt = 0.125 * 2.54;
            delta = dt / N;
            for(int i = 1; i <= level; i++)
            {
                dimention1[i] = Math.Sqrt(((Length[i] / 30.48)* (Length[i] / 30.48)) + (((H[i] - H[i - 1] - delta) / 30.48)* ((H[i] - H[i - 1] - delta) / 30.48)));
                dimention2[i] = (Length[i] / 30.48) / 3;
            }

            string[] oTitles = new string[Row];

            oTitles[0] = "MARK";
            oTitles[1] = "DIMENTION A";
            oTitles[2] = "DIMENTION B";
            oTitles[3] = "BOTTOM EDGE";
            oTitles[4] = "VERTICAL EDGE";

            string[] oContents=new string[Row*level];
            function obj = new function();
            int k = 0;
            int j = 1;
            for(int i=0;i< Row * level; i++)
            {
                if(i % Row == 0) {
                    k = k + 1;
                    oContents[i] = "SR" + k;
                //'oContents(i) = "k"
               }


                 if(i % Row == 1) 
                {
                    obj.toftin(dimention1[k], ref ar);
                   
                 if (((double)ar[3] / (double)ar[4]) <= 0.0625) {
                        oContents[i] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + "''";
                    }
                 else {
                        oContents[i] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + " " + Convert.ToString(ar[3]) + "/" + Convert.ToString(ar[4]) + "''";
                    }
                  //' k = k + 1
                //'oContents(i) = "k"
                 }


                    if(i % Row == 2) {
                    obj.toftin(dimention2[k], ref ar);

                        if (((double)ar[3] / (double)ar[4]) <= 0.0625)
                        {
                            oContents[i] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + "''";
                        }
                        else { 
                            oContents[i] = Convert.ToString(ar[1]) + "'-" + Convert.ToString(ar[2]) + " " + Convert.ToString(ar[3]) + "/" + Convert.ToString(ar[4]) + "''";
                    }
                   }


                     if(i % Row == 3) {
                                    oContents[i] = "SQUARE";
                   }


                     if(i % Row == 4) {
                                    oContents[i] = "SQUARE";
                   }



            }

            CustomTable oCustomTable;

            oCustomTable = oSheet.CustomTables.Add("SHELL PLATE TABLE", ThisApplication.TransientGeometry.CreatePoint2d(1 * 2.54, 3.5 * 2.54), Row, level, oTitles, oContents);
            //' Change the 3rd column to be left justified.
            oCustomTable.Columns[3].ValueHorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
            oDrawDoc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.TableStyle.RowGap = 0.1;

            TableFormat oFormat;
            oFormat = oSheet.CustomTables.CreateTableFormat();
            TextStyle oTextstyles;
            oTextstyles = oDrawDoc.StylesManager.TextStyles[18];
            oTextstyles.Font = "RomanS";
            oTextstyles.FontSize = 0.15;
            oFormat.TextStyle = oTextstyles;

            oFormat.OutsideLineWeight = 0.05;
            oCustomTable.OverrideFormat = oFormat;
        }

        public void CRtable4(Inventor.Application ThisApplication,int level,Double N,Double[,] coord,Double[] H,Double[] Length,  String name, Double[] Thickness, String[] sdiscription, String[] material,String[] note, Double Radius, Double alfa) 
        {
            DrawingDocument oDrawDoc;
            oDrawDoc = (DrawingDocument)ThisApplication.ActiveDocument;
            int[] ar = new int[5];
            Sheet oSheet;
            oSheet = oDrawDoc.Sheets[3];
            int Row;
            Row = 9;
            double Total;
            function obj = new function();
            string[] oTitles = new string[Row];
            oTitles[0] = "SA/PC";
            oTitles[1] = "MK";
            oTitles[2] = "QTY";
            oTitles[3] = "SHIP DESCRIPTION";
            oTitles[4] = "DESCRIPTION";
            oTitles[5] = "LGTH";
            oTitles[6] = "MATL";
            oTitles[7] = "Notes";
            oTitles[8] = "WT/EA";

            Total = 0;

            for (int i = 1; i <= level; i++) {
                Total = Total + ((Length[i] / 30.48) * (Thickness[i] / 2.54) * ((H[i] - H[i - 1]) / 2.54) * 0.2836 * 12) * N;
                }

            string[] oContents = new string[Row];
           // function obj = new function();
            int k = 0;
           // int j = 1;
           for(int i = 0; i < Row; i++)
            {
                if(i % Row == 0) {
                    k = k + 1;
                    oContents[i] = Convert.ToString(k);

                }


                if(i % Row == 1) {
                    oContents[i] = "SR" + k;
    
                }


                if (i % Row == 2)
                {
                    oContents[i] = Convert.ToString(N);
                }

                if(i % Row == 3) {
                    oContents[i] = "TANK SHELL ERECTION";
               // 'oContents(i) = "k"
                }


                if(i % Row == 4) {
                        //'oContents(i) = "PL " + (Thickness(k)) / 2.54 + " in x " + (H(k) - H(k - 1)) / 2.54 + "''"
                   oContents[i] = " ";

                }


                 if(i % Row == 5) {
                   oContents[i] = " ";
                }


                if(i % Row == 6) {

                   oContents[i] = " ";
                }


                 if(i % Row == 7) {

                   oContents[i] = " ";
                }


                 if(i % Row == 8) {

                  oContents[i] = Convert.ToString(Math.Round(Total, 2));
                }

            }


            CustomTable oCustomTable;

            oCustomTable = oSheet.CustomTables.Add("BILL OF MATERIAL", ThisApplication.TransientGeometry.CreatePoint2d(6 * 2.54, 10 * 2.54), Row, 1, oTitles, oContents);
            //' Change the 3rd column to be left justified.
            Double[] sheetsize = new double[3];
            sheetsize[1] = 13.38 * 2.54;
            sheetsize[2] = 8.1875 * 2.54;

            Double Oheight;
            Oheight = (oCustomTable.RangeBox.MaxPoint.Y - oCustomTable.RangeBox.MinPoint.Y) / 2.54;



            Double oLength;
            oLength = (oCustomTable.RangeBox.MaxPoint.X - oCustomTable.RangeBox.MinPoint.X) / 2.54;


            oCustomTable.Position = ThisApplication.TransientGeometry.CreatePoint2d(sheetsize[1] - oLength * 2.54 - 0.394 * 2.54, sheetsize[2] - 0.394 * 2.54);
            oCustomTable.Columns[3].ValueHorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
            oDrawDoc.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.TableStyle.RowGap = 0.1;

            TableFormat oFormat;
            oFormat = oSheet.CustomTables.CreateTableFormat();
            TextStyle oTextstyles;
            oTextstyles = oDrawDoc.StylesManager.TextStyles[18];
            oTextstyles.Font = "RomanS";
            oTextstyles.FontSize = 0.15;
            oFormat.TextStyle = oTextstyles;

            oFormat.OutsideLineWeight = 0.05;
            oCustomTable.OverrideFormat = oFormat;


        }
    }
}
