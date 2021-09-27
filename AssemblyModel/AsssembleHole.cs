using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace AssemblyModel
{
    class AssembleHole
    {
        public AssembleHole(Inventor.Application InventorApplication, AssemblyDocument oAssyDoc, object[] A,ComponentOccurrence oC2)
        {
           // AssemblyDocument oAssyDoc;
           // oAssyDoc = (AssemblyDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);
            TransientGeometry oTransGeom = default(TransientGeometry);
            oTransGeom = InventorApplication.TransientGeometry;

            Matrix oPositionMatrix, oPositionMatrix2;

            oPositionMatrix = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
            oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();

            string sfilename;
            sfilename = "C:\\Rahul\\Nozzle\\NozzleAssembly.iam";
            ComponentOccurrence oC1;
            oC1 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);

            //ComponentOccurrence oC1_4;
            //int lastcomponentcount;

           
            //lastcomponentcount = oC1.SubOccurrences.Count;
           // oC1.SubOccurrences[lastcomponentcount].Visible = false;
           //  oC1.SubOccurrences[lastcomponentcount].Enabled = false;
            //lastcomponentcount = oC1.SurfaceBodies.Count;
            //oC1.SurfaceBodies[lastcomponentcount].Visible = false;
            //oC1_4 = oC1.ComponentDefinition.Occurrences[1];
            //oView1.SetVisibility(oC, false);

            AssemblyComponentDefinition pcd;
            pcd = (AssemblyComponentDefinition)oC1.Definition;
            Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
            oPositionMatrix.SetTranslation(oTrans);
            WorkPlane oworkplane1;
            oworkplane1 = pcd.WorkPlanes[1];
            WorkAxis oworkaxis1;
            oworkaxis1 = pcd.WorkAxes[1];
            WorkAxis oworkaxis2;
            oworkaxis2 = pcd.WorkAxes[2];

            functions obj2 = new functions();
            obj2.WorkPlaneC(oAssyDoc, oC1, oC2, oworkplane1,(WorkPlane)A[0]);
            obj2.AxisContraints(oAssyDoc, oC1, oC2, oworkaxis1, (WorkAxis)A[1]);
            obj2.AxisContraints(oAssyDoc, oC1, oC2, oworkaxis2, (WorkAxis)A[2]);

            //oAssyDoc.SaveAs("C:\\Rahul\\Shell\\FinalAssembly.iam",false);
            InventorApplication.SilentOperation = true;
            oAssyDoc.Save();
            InventorApplication.SilentOperation = false;
            //string sfilename2="abc";
        }
    }
}
