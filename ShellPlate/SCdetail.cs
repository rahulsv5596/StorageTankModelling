using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShellPlate
{
   
    public partial class SCdetail : Form
    {
        List<TextBox> TextBoxList = new List<TextBox>();
       

        int Row2;
        public SCdetail(int Row,int Col)
        {
            Row2 = Row;
            int A = 1, B = 1;
           
            InitializeComponent();
            for(int i = 1; i <= Row; i++)
            {
                A = A + 1;
                B = 15;
                for (int j = 1; j <= 4; j++) 
                {

                    //System.Windows.Forms.TextBox Txt2 = new System.Windows.Forms.TextBox();
                    AddNewTextBox(Row, Col, A, B);
                    
                   // TextBoxList.Add(Txt2);
                    B = B + 105;
                }
                
            }

            //int k=0;
            //for(int i = 1; i <= Row; i++)
            //{
            //    for(int j = 1; j <= 4; j++)
            //    {
            //        Console.WriteLine(TextBoxList[k].Text);
            //        k = k + 1;
            //    }
            //}
          
            
        }

        private void SCdetail_Load(object sender, EventArgs e)
        {

        }
        int i=0, B;
        public  System.Windows.Forms.TextBox AddNewTextBox(int Row,int Col,int A,int B)
        {
            System.Windows.Forms.TextBox Txt = new System.Windows.Forms.TextBox();

            TextBoxList.Add(Txt);

            this.Controls.Add(TextBoxList[i]);
            TextBoxList[i].Name = "TextBox" + i.ToString();
            TextBoxList[i].Top = 28*A;
            TextBoxList[i].Left = B;
            //TextBoxList[i].Text = Txt.Name;

            i = i + 1;
            return TextBoxList[i-1];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Update();
            this.Close();
            
        }

        public List<TextBox> GetList()
        {
            return TextBoxList;
        }

    }
}
