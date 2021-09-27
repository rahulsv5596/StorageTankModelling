using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using ShellPlate;
using Microsoft.VisualBasic;

namespace AssemblyModel
{
    public partial class AssembleN : Form
    {
        bool flag = false;
        Inventor.Application InventorApplication;
        public AssembleN()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            flag = true;

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            flag = false;
        }
        private void assembleB_Click(object sender, EventArgs e)
        {
            try
            {
                InventorApplication = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                //MessageBox.Show("Yahoo inventor connected succesfully");

            }
            catch
            {
                Type invType = System.Type.GetTypeFromProgID("Inventor.Application");
                InventorApplication = System.Activator.CreateInstance(invType) as Inventor.Application;
                InventorApplication.Visible = true;
            }

            //ShellPlate.Shell obj = new ShellPlate.Shell();



            AssemblyDocument oAssyDoc;
            oAssyDoc = (AssemblyDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);
            InventorApplication.SilentOperation = true;
            oAssyDoc.SaveAs("C:\\Rahul\\Shell\\FinalAssembly.iam", false);
            InventorApplication.SilentOperation = false;
            TransientGeometry oTransGeom = default(TransientGeometry);
            oTransGeom = InventorApplication.TransientGeometry;
            //Shell component
            Matrix oPositionMatrix;

            oPositionMatrix = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
            // oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();

            string sfilename;
            sfilename = "C:\\Rahul\\shell\\CBA.ipt";
            ComponentOccurrence oC1;
            oC1 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix);

            Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
            oPositionMatrix.SetTranslation(oTrans);


            //Vector oTrans;
            oTrans = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
            oPositionMatrix.SetTranslation(oTrans);
            SurfaceBody oBody;
            string f = "C:\\Rahul\\";
            if (flag == true)
            {
                //MessageBox.Show("Nozzle Will Not be in consideration");

                MessageBox.Show("Nozzle Will be in consideration");
                int  N;
                N=Convert.ToInt32(Interaction.InputBox("Number of Nozzles", "Nozzle Number"));
                Double iangle=0;
                //oBody=oC1.SurfaceBodies[1];
                //Face oface;
                //oface = getMaxface(oBody);
                string[] iangles;
                iangles = (Interaction.InputBox("Enter the angle in radian Clockwise", "Nozzle Number","1,2,3.....")).Split(',');
               
                double[] iangleValues=new double[N+1];
                int k = 1;
                foreach(string i in iangles)
                {
                    iangleValues[k] = Convert.ToDouble(i);
                    k++;
                }
                iangleValues[0]=0;
                for (int i = 0; i < N; i++) {
                  
                    Hole obj2 = new Hole();
                    // ShellPlate.Shell o
                    obj2.CreateHole(f, InventorApplication, oAssyDoc,iangleValues[i]);
                }
               
                // AssemblyDocument oAssyDoc;
                oAssyDoc = (AssemblyDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);


 

                //assembly Component
                Matrix oPositionMatrix3;

                oPositionMatrix3 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
                // oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();

                string sfilename3;
                
                sfilename = "C:\\Rahul\\Shell\\FinalAssembly.iam";
                
                ComponentOccurrence oC3;
                oC3 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix3);

                Vector oTrans3;
                oTrans3 = InventorApplication.TransientGeometry.CreateVector(1, 0, 0);
                oPositionMatrix.SetTranslation(oTrans3);

                //Base Component
                Matrix oPositionMatrix2;

                oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
                // oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();

                string sfilename2;
                sfilename = "C:\\Rahul\\Base\\base.ipt";
                ComponentOccurrence oC2;
                oC2 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix2);

                functions custom = new functions();

              
                functions A_C = new functions();
                AssemblyComponentDefinition pcd1;
                PartComponentDefinition pcd2;
                pcd1 = (AssemblyComponentDefinition)oC3.Definition;
                WorkAxis oworkaxis1;
                oworkaxis1 = pcd1.WorkAxes[2];
                pcd2 = (PartComponentDefinition)oC2.Definition;
                WorkAxis oworkaxis2;
                oworkaxis2 = pcd2.WorkAxes[2];

                //A_C.FaceContraints(oAssyDoc, oC1, oC2A, 1, 8);
                A_C.AxisContraints(oAssyDoc, oC3, oC2, oworkaxis1, oworkaxis2);
                double width = Convert.ToDouble(custom.ReadCustomData("width", f + "shell\\CBA.ipt", InventorApplication));
                A_C.CenterContstraints2(oAssyDoc, oC3, oC2, width * 2.54);


            }
            else
            {
                MessageBox.Show("Nozzle Will Not be in consideration");
                //Base Component
                Matrix oPositionMatrix2;

                oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();
                // oPositionMatrix2 = (Matrix)InventorApplication.TransientGeometry.CreateMatrix();

                string sfilename2;
                sfilename = "C:\\Rahul\\Base\\base.ipt";
                ComponentOccurrence oC2;
                oC2 = oAssyDoc.ComponentDefinition.Occurrences.Add(sfilename, oPositionMatrix2);
                functions custom = new functions();

                
                functions A_C = new functions();
                PartComponentDefinition pcd1,pcd2;
                pcd1 = (PartComponentDefinition)oC1.Definition;
                WorkAxis oworkaxis1;
                oworkaxis1 = pcd1.WorkAxes[2];
                pcd2 = (PartComponentDefinition)oC2.Definition;
                WorkAxis oworkaxis2;
                oworkaxis2 = pcd2.WorkAxes[2];

                //A_C.FaceContraints(oAssyDoc, oC1, oC2A, 1, 8);
                A_C.AxisContraints(oAssyDoc, oC1, oC2, oworkaxis1, oworkaxis2);
                double width = Convert.ToDouble(custom.ReadCustomData("width", f + "shell\\CBA.ipt", InventorApplication));
                A_C.CenterContstraints(oAssyDoc, oC1, oC2, width*2.54);

                
                MessageBox.Show("finish Line Crossed");
            }

            ObjectVisibility objv=default(ObjectVisibility);
            objv = oAssyDoc.ObjectVisibility;
            objv.AllWorkFeatures=false;

            oAssyDoc.SaveAs("C:\\Rahul\\Shell\\FinalAssembly2.iam", false);
            // Appearance ox;
            //Apearence ox;


        }

        
         
      
    }
}
