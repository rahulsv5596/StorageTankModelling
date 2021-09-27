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


namespace ShellPlate
{
    public partial class Shell : Form
    {
        List<System.Windows.Forms.TextBox> TextBoxList = new List<System.Windows.Forms.TextBox>();
        Inventor.Application InventorApplication;
        String f= "C:\\Rahul\\shell\\";
      

        public Shell()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox5.BackColor == System.Drawing.Color.Red)
            {
                MessageBox.Show("Change the Number of Devision");
            }

            else
            {
                int Row = Convert.ToInt32(textBox2.Text);
                SCdetail obj = new SCdetail(Convert.ToInt32(textBox2.Text),Convert.ToInt32(textBox4.Text));

                obj.ShowDialog();
                //this.Close();

                TextBoxList = obj.GetList();
                Console.WriteLine(TextBoxList.Count);
                //for(int i = 0; i < 10; i++)
                //{
                
            }

            
        }


        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (Convert.ToDouble(textBox5.Text) < 40)
            {
               textBox5.BackColor = System.Drawing.Color.GreenYellow;
            }
            else
            {
                textBox5.BackColor = System.Drawing.Color.Red;
            }
           
        }



        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.TextLength != 0 && textBox2.TextLength != 0 && textBox3.TextLength != 0)
            {
                //textBox5.BackColor = Color.Green;
                textBox5.Text = (Math.Round(3.14 * Convert.ToDouble(textBox1.Text) / Convert.ToDouble(textBox3.Text),2)).ToString();
            }
            else
            {
                textBox5.BackColor = System.Drawing.Color.White;
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                InventorApplication = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                MessageBox.Show("Yahoo inventor connected succesfully");

            }
            catch
            {
                Type invType = System.Type.GetTypeFromProgID("Inventor.Application");
                InventorApplication = System.Activator.CreateInstance(invType) as Inventor.Application;
                InventorApplication.Visible = true;
            }
            int k = 0;
            int level;
            level = Convert.ToInt32(textBox2.Text);
            double radius;
            radius = Convert.ToDouble(textBox1.Text) * 30.48 / 2;
            int N;
            N = Convert.ToInt32(textBox3.Text);
            double[] H = new double[level+1];

            double dt;
            dt = Convert.ToDouble(textBox6.Text) * 2.54*N;
            double[] Thickness = new double[level + 1];
            double[] Height = new double[level + 1];
            string[] Course = new string[level + 1];
            string[] material = new string[level + 1];
            Console.WriteLine("in the shell");

            k = 1;
            for (int j = 0; j < TextBoxList.Count; j++)
            {
                //System.Windows.Forms.TextBox Txt = new System.Windows.Forms.TextBox();
                //txt = this.Controls.Contains(TextBoxList[k]);
                //Console.WriteLine("Modified");
                Console.WriteLine(TextBoxList[j].Text);

                if ((j % 4) == 0)
                {
                    Course[k] = TextBoxList[j].Text;
                }
                if ((j % 4) == 1)
                {
                    material[k] = TextBoxList[j].Text;
                }
                if ((j % 4) == 2)
                {
                    Height[k] = Convert.ToDouble(TextBoxList[j].Text);
                    H[k] = H[k - 1] + Height[k] * 2.54;

                }
                if ((j % 4) == 3)
                {
                    Thickness[k] = Convert.ToDouble(TextBoxList[j].Text)*2.54;
                    k = k + 1;
                }
                
                if (k > level)
                {
                    break;
                }
            }

            string[] note = new string[level+1];
            string[] sdiscription = new string[level+1];

            
            shell_2d shell2dobj = new shell_2d();

            for (int i = 1; i <= level; i++)
            {
                note[i] = "NOTE-4.4";
                sdiscription[i] = "SHELL PLATE";
            }


         
            shell2dobj.Shell2d(InventorApplication, level, radius, N, H, Thickness, dt, material, note, sdiscription);
            
            MessageBox.Show("Model is Created");
            this.Close();
        }

        private void Shell_Load(object sender, EventArgs e)
        {

        }
    }
}
