using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsnew;
using ShellPlate;
using BottomPlate;
using System.Windows;
using AssemblyModel;

namespace InvAddIn
{
    public partial class InitForm : Form
    {
        public InitForm()
        {
            InitializeComponent();
        }

        private void InitForm_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            WindowsFormsnew.Nozzle objN = new WindowsFormsnew.Nozzle();
            objN.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShellPlate.Shell objS = new ShellPlate.Shell();
            objS.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            BottomPlate.BOTP objB = new BottomPlate.BOTP();

            objB.Show();

        }

        private void AssembleButton_Click(object sender, EventArgs e)
        {
            AssemblyModel.AssembleN objB = new AssemblyModel.AssembleN();

            objB.Show();

        }
    }
}
