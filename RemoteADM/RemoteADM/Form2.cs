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
        int i = 0,z=0,p=0;
        public static Form2 ss;
       public static string slogans;
        string[] breakMysentence;
        public Form2()
        {
            InitializeComponent();
          
        }
        void sloganchecker()
        {
            foreach(char c in slogans)
            {
                if (c == '.')
                {
                    i++;
                }
               
            }
            breakMysentence = slogans.Split('.');

        }
        public void Form2_Load(object sender, EventArgs e)
        {
            DataStorage data = new DataStorage();
            slogans = data.sloganString();
            sloganchecker();
       
            timer3.Enabled = true;
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
            z++;
            if (z == 1000)
            {
                z = 0;
                p++;
            }
            if (p ==(i))
            {
                
                    p = 0;
                
            }
         
               
                label_Slogan.Text = breakMysentence[p];
              
            
              
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void label_Slogan_Click(object sender, EventArgs e)
        {

        }
    }
}
