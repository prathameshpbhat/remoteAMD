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
        int i = 0,z=0,p=0,j=0;
        public static Form2 ss;
       public static string slogans;
        string[] breakMysentence;
        public Form2()
        {
            InitializeComponent();
          
        }
        void sloganchecker()
        {
            slogans = "हमेशा सतर्क रहिये और दुर्घटना रोकीये| .जीवन बहुमूल्य है, संरक्षा ध्यान में रखें। . संरक्षपूर्वक काम करनेकी आदत डालें।.थोड़ीसी सावधानी दुर्घटना कोदुर्लभ बना देती है।.संरक्षा को अवकाश नहीं है, लापरवाही से जाने जाती है। . सुरक्षा नियमों का हमेशा पालन करें।. एक सतर्क लोको पायलट ही संरक्षा का सबसे बेहतर साधन है| . सतर्क रहे, सुरक्षित रहे। सभी बीमारियों की दवा स्वच्छता ही है।.रेल हमारी शान है, सफाई हमारी पहचान है। . Even while at work,play for safety.";
            foreach (char c in slogans)
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
  
            label_Slogan.Text = "";

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

            if (j == i)
                j = 0;
            label_Slogan.Text = breakMysentence[j];
            j++;


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint_2(object sender, PaintEventArgs e)
        {

        }

        private void label_Date_Click(object sender, EventArgs e)
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
