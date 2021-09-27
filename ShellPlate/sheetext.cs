using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace ShellPlate
{
    class sheetext
    {
        public void SheetTextAdd1(Inventor.Application InventorApplication,string f,double[] sheetsize)
        {
            DrawingDocument oDrawdoc;
            oDrawdoc = (DrawingDocument)InventorApplication.Documents.Open(f+"Shelldwg.dwg",true);
            Sheet oActiveSheet;
            oActiveSheet = oDrawdoc.Sheets[1];
            GeneralNotes oGeneralNotes;
            oGeneralNotes = oActiveSheet.DrawingNotes.GeneralNotes;
            TransientGeometry oTG;
            oTG = InventorApplication.TransientGeometry;

            String sText;
            Double dYcoord, CN;
            CN = sheetsize[2] / 2.54 - 0.4875;
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
            sText = "WPG-AS NOTED";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "2.");
            sText = "NDE-RT1,SEE QA/QC SHELL RADIOGRAPH";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x1 * 2.54, dYcoord), "3.");
            sText = "DWG UNLESS NOTED";
            oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);
            //dYcoord = dYcoord - (oGeneralNote.FittedTextHeight + gap);
            //sText = "HEAVY GREASE AFTER TESTING";
            //oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x2 * 2.54, dYcoord), sText);

        }
    }
}
