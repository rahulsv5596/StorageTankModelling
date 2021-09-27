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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsnew
{
    public partial class Nozzle : Form
    {
        Inventor.Application InventorApplication;
        public Nozzle()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender,EventArgs e)
        {
            CB_1.Items.Add(new KeyValuePair<string, string>("YES", "0"));
            CB_1.Items.Add(new KeyValuePair<string, string>("NO", "1"));
            CB_1.DisplayMember = "key";
            CB_1.ValueMember = "value";

            PT.Items.Add(new KeyValuePair<string, string>("Standard" , "0"));
            PT.Items.Add(new KeyValuePair<string, string>("Extra Strong", "1"));
            PT.DisplayMember = "key";
            PT.ValueMember = "value";

            NT.Items.Add(new KeyValuePair<string, string>("WeldneckNozzle", "0"));
            NT.Items.Add(new KeyValuePair<string, string>("SlipOnNozzle", "1"));
            NT.DisplayMember = "key";
            NT.ValueMember = "value";
        }



        private void Button1_Click(object sender, EventArgs e)
        {
          


            MessageBox.Show("This Will take few minutes");
            functions EX_read = new functions();

            double[] textboxes = new double[7];


            textboxes[1] = Convert.ToDouble(textBox1.Text) * 30.48;
            textboxes[2] = Convert.ToDouble(textBox2.Text) * 2.54;
            textboxes[3] = Convert.ToDouble(textBox3.Text) * 2.54;
            textboxes[4] = Convert.ToDouble(textBox4.Text) * 2.54;
            textboxes[5] = 2.25 * textboxes[3];
            textboxes[6] = 2.25 * textboxes[3];

            double C_Value = textboxes[4]/2.54;
            double[] Narr = new double[10];
            double[] Farr = new double[12];
            string Path= "C:\\Rahul\\Nozzle\\Dimensions of Shell nozzle.xlsx";
            EX_read.ExcelRead(C_Value, ref Narr, Path,30,5);
            // MessageBox.Show((C_Index).ToString());
            int i = 0;
            for (i = 1; i <= 5; i++)
            {
                //arr[j] = Convert.ToDouble((excel.ReadCell(C_Index, j)));
                Console.WriteLine("array values {0}", Narr[i]);
            }


            Path = "C:\\Rahul\\Nozzle\\nozzle Parameter.xlsx";
            if (C_Value <= 24)
            {
                EX_read.ExcelRead(C_Value, ref Farr, Path, 21, 11);
            }
            textboxes[6] = 1.25*Narr[2];

            Path = "C:\\Rahul\\Nozzle\\Pipe Weights.xlsx";
            double[] PipeWeights = new double[4];
            //MessageBox.Show(Convert.ToString(CB_1.Text));
            EX_read.ExcelRead(C_Value, ref PipeWeights, Path, 18, 3);
            double NozzleWeight;

            if(PT.Text=="Extra Strong")
            {
                NozzleWeight = PipeWeights[3]/2.54;
            }
            else
            {
                NozzleWeight = PipeWeights[2]/2.54;
            }


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

            if (NT.Text == "WeldneckNozzle")
            {
                String[] Ship_Discription = new string[5];
                Ship_Discription[0] = "NOZZLE";
                Ship_Discription[1] = "REPAD";

                string[] MK = new string[5];
                MK[0] = "A1";
                MK[1] = "A1R";
                MK[2] = "A1F";


                String[] Discription = new string[5];
                Discription[0] = "SHELL NOZZLE ASSEMBLY-B";
                Discription[1] = "PL " + textBox2.Text + " in X" + Convert.ToString(Narr[2] / 2.54) + "''";


                String[] MATL = new string[5];
                MATL[0] = "";
                MATL[1] = "G40.21-38W";


                double lengthPS, lengthPN;


                lengthPS = Narr[3];
                lengthPN = Narr[3];

                string CB;
                // CB = Convert.ToString(CB_1);
                Double[] LGTH = new double[5];
                LGTH[3] = Math.Abs(lengthPN + lengthPS);

                Component_1 Cylender = new Component_1(InventorApplication, Narr, Farr, CB_1.Text);
                Component_2 Nozzle = new Component_2(InventorApplication, Narr, Farr, CB_1.Text);
                if ((textboxes[3] / 2.54) <= (Narr[4] / 2.54))
                {
                    Component_3 Shield = new Component_3(InventorApplication, Narr, Farr, textboxes);
                    LGTH[1] = Narr[2] / 2 + textboxes[3];

                }
                else
                {
                    Component_3_b Shield = new Component_3_b(InventorApplication, Narr, Farr, textboxes);
                    LGTH[1] = Narr[2];
                }

                ShellComponent oShell = new ShellComponent(InventorApplication, Narr, Farr, textboxes);
                Assembly ASM = new Assembly(InventorApplication, Narr, Farr, CB_1.Text);
                Assemblyflat ASM2 = new Assemblyflat(InventorApplication, Narr, Farr, CB_1.Text);



                double length;
                double height;
                height = textboxes[6];
                length = textboxes[5];



                Double[] Varr = new Double[3];
                string f = "C:\\Rahul\\Nozzle\\C1.ipt";
                functions V = new functions();
                Varr[0] = V.VolumeC(InventorApplication, f, 1);

                f = "C:\\Rahul\\Nozzle\\C2.ipt";
                Varr[1] = V.VolumeC(InventorApplication, f, 1);

                f = "C:\\Rahul\\Nozzle\\C3.ipt";
                Varr[2] = V.MAx_Area(InventorApplication, f, 1);
                Console.WriteLine("array value V[0] : {0},V[0] : {1},V[0] : {2}", Varr[0], Varr[1], Varr[2]);

                double[] weight = new double[5];

                weight[0] = Math.Round((LGTH[3] / (2.54 * 12)) * NozzleWeight + Farr[9] / 2.54, 2);
                weight[1] = Math.Round((Varr[2] / (2.54 * 2.54 * 2.54)) * textboxes[2], 0);
                if (((Varr[2] / (2.54 * 2.54 * 2.54)) * (textboxes[2]) - Math.Round((Varr[2] / (2.54 * 2.54 * 2.54)) * textboxes[2], 0)) < 0)
                {
                    weight[1] = Math.Round(weight[1] * 0.283555556, 2);
                }
                else
                {
                    weight[1] = Math.Round((weight[1] + 1) * 0.283555556, 2);
                }


                double[] qty = new double[4];
                qty[0] = 1;
                qty[1] = 1;
                qty[2] = 1;
                qty[3] = 1;

                weight[2] = Math.Round(Farr[9] / 2.54, 2);
                weight[3] = Math.Round((LGTH[3] / (2.54 * 12)) * NozzleWeight, 2);
                weight[4] = Math.Round(Farr[9] / 2.54, 2);

                Ship_Discription[2] = "INTERNAL FLANGE";
                Ship_Discription[3] = "PIPE " + textBox4.Text + "''XS";
                Ship_Discription[4] = "FLANGE " + textBox4.Text + "'' 150# RFWN XS BORE";

                Discription[2] = "FLANGE " + textBox4.Text + "'' 150# RFWN XS BORE";
                Discription[3] = "PIPE " + textBox4.Text + "'' XS";
                Discription[4] = "FLANGE " + textBox4.Text + "'' 150# RFWN XS BORE";

                MATL[2] = "A105";
                MATL[3] = "A106";
                MATL[4] = "A105";
                if (CB_1.Text == "YES")
                {

                    // weight[3] = weight[3];

                    qty[3] = 2;


                }

                for (i = 0; i <= 3; i++)
                {
                    Console.WriteLine("weight {0}:{1}", i, weight[i]);
                }

                Double Roll_radius;
                Roll_radius = (Convert.ToDouble(textBox1.Text) * 12 + Convert.ToDouble(textBox2.Text)) * 2.54;
                Drawing_Document oDoc = new Drawing_Document(InventorApplication, length, height, Ship_Discription, Discription, MATL, LGTH, weight, qty, Roll_radius, MK);
                MessageBox.Show("Modelling Finished");
            }
            else
            {

                String[] Ship_Discription = new string[5];
                Ship_Discription[0] = "NOZZLE";
                Ship_Discription[1] = "REPAD";

                string[] MK = new string[5];
                MK[0] = "A1";
                MK[1] = "A1R";
                MK[2] = "A1F";


                String[] Discription = new string[5];
                Discription[0] = "SHELL NOZZLE ASSEMBLY-B";
                Discription[1] = "PL " + textBox2.Text + " in X" + Convert.ToString(Narr[2] / 2.54) + "''";


                String[] MATL = new string[5];
                MATL[0] = "";
                MATL[1] = "G40.21-38W";


                double lengthPS, lengthPN;

                //if (CB_1.Text == "NO")
                //{
                //    lengthPS = Narr[3] - Farr[6];
                //    lengthPN = Narr[3];
                //    //Console.WriteLine("here we go");
                //}
                //else
                //{
                //    lengthPS = Narr[3] - Farr[6];
                //    lengthPN = Narr[3] - Farr[6];

                //}
                lengthPS = Narr[3];
                lengthPN = Narr[3];

                string CB;
                // CB = Convert.ToString(CB_1);
                Double[] LGTH = new double[5];
                LGTH[3] = Math.Abs(lengthPN + lengthPS);

                Component_1B2 Cylender = new Component_1B2(InventorApplication, Narr, Farr, CB_1.Text);
                Component_2B Cylender2 = new Component_2B(InventorApplication, Narr, Farr, CB_1.Text);

                if ((textboxes[3] / 2.54) <= (Narr[4] / 2.54))
                {
                    Component_3B2 Repad = new Component_3B2(InventorApplication, Narr, Farr, textboxes);
                    //Component_3 Shield = new Component_3(InventorApplication, Narr, Farr, textboxes);
                    LGTH[1] = Narr[2] / 2 + textboxes[3];

                }
                else
                {
                    Component_3B2_C Repad = new Component_3B2_C(InventorApplication, Narr, Farr, textboxes);
                    //Component_3_b Shield = new Component_3_b(InventorApplication, Narr, Farr, textboxes);
                    LGTH[1] = Narr[2];
                }

                shellcomponent_B oShell = new shellcomponent_B(InventorApplication, Narr, Farr, textboxes);
                // Console.WriteLine("CB1.text: {0]", CB_1.Text);
                Assembly2B ASM = new Assembly2B(InventorApplication, Narr, Farr, CB_1.Text);
                Assemblyflat ASM2 = new Assemblyflat(InventorApplication, Narr, Farr, CB_1.Text);
                double length;
                double height;
                height = textboxes[6];
                length = textboxes[5];



                Double[] Varr = new Double[3];
                string f = "C:\\Rahul\\Nozzle\\C1.ipt";
                functions V = new functions();
                Varr[0] = V.VolumeC(InventorApplication, f, 1);

                f = "C:\\Rahul\\Nozzle\\C2.ipt";
                Varr[1] = V.VolumeC(InventorApplication, f, 1);

                f = "C:\\Rahul\\Nozzle\\C3.ipt";
                Varr[2] = V.MAx_Area(InventorApplication, f, 1);
                Console.WriteLine("array value V[0] : {0},V[0] : {1},V[0] : {2}", Varr[0], Varr[1], Varr[2]);

                double[] weight = new double[5];

                weight[0] = Math.Round((LGTH[3] / (2.54 * 12)) * NozzleWeight + Farr[11] / 2.54, 2);
                weight[1] = Math.Round((Varr[2] / (2.54 * 2.54 * 2.54)) * textboxes[2], 0);
                if (((Varr[2] / (2.54 * 2.54 * 2.54)) * (textboxes[2]) - Math.Round((Varr[2] / (2.54 * 2.54 * 2.54)) * textboxes[2], 0)) < 0)
                {
                    weight[1] = Math.Round(weight[1] * 0.283555556, 2);
                }
                else
                {
                    weight[1] = Math.Round((weight[1] + 1) * 0.283555556, 2);
                }


                double[] qty = new double[4];
                qty[0] = 1;
                qty[1] = 1;
                qty[2] = 1;
                qty[3] = 1;

                weight[2] = Math.Round(Farr[11] / 2.54, 2);
                weight[3] = Math.Round((LGTH[3] / (2.54 * 12)) * NozzleWeight, 2);
                weight[4] = Math.Round(Farr[11] / 2.54, 2);

                Ship_Discription[2] = "INTERNAL FLANGE";
                Ship_Discription[3] = "PIPE " + textBox4.Text + "''XS";
                Ship_Discription[4] = "FLANGE " + textBox4.Text + "'' 150# RFWN XS BORE";

                Discription[2] = "FLANGE " + textBox4.Text + "'' 150# RFWN XS BORE";
                Discription[3] = "PIPE " + textBox4.Text + "'' XS";
                Discription[4] = "FLANGE " + textBox4.Text + "'' 150# RFWN XS BORE";

                MATL[2] = "A105";
                MATL[3] = "A106";
                MATL[4] = "A105";
                if (CB_1.Text == "YES")
                {

                    // weight[3] = weight[3];

                    qty[3] = 2;


                }

                for (i = 0; i <= 3; i++)
                {
                    Console.WriteLine("weight {0}:{1}", i, weight[i]);
                }

                Double Roll_radius;
                Roll_radius = (Convert.ToDouble(textBox1.Text) * 12 + Convert.ToDouble(textBox2.Text)) * 2.54;
                Drawing_DOC2 oDoc = new Drawing_DOC2(InventorApplication, length, height, Ship_Discription, Discription, MATL, LGTH, weight, qty, Roll_radius, MK);
                MessageBox.Show("Modelling Finished");
                // Drawing_Document oDoc = new Drawing_Document(InventorApplication, length, height, Ship_Discription, Discription, MATL, LGTH, weight, qty, Roll_radius, MK);
            }


        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
