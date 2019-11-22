using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RemoteADM
{
   
    public partial class Form2 : Form
    {
        int i = 0;
        public static Form2 ss;
        public Form2()
        {
            InitializeComponent();
          
        }

        public void Form2_Load(object sender, EventArgs e)
        {
            label_Slogan.Text = "Don't be safely blinded be safety minded.";
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if(i==0)
            {
                label_Slogan.ForeColor = Color.Green;
                label_Slogan.Text = "Don't be safely blinded be safety minded.";
                i = 1;
            }
            else
            {
                label_Slogan.ForeColor = Color.Blue;
                label_Slogan.Text = "Being safe is like Breathing you never want to stop.";
                i = 0;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
