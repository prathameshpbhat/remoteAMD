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
    public partial class settingform : Form
    {
        System.Drawing.Color Header,Secondheader;
        int b1=0, b2=0;
        DataStorage dataStorage=new DataStorage();
        colorchange c=new colorchange();
        public settingform()
        {
            InitializeComponent();
        }

        private void settingform_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                Header = colorDlg.Color;
                button1.BackColor = colorDlg.Color;
                b1 = 1;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                Secondheader = colorDlg.Color;
                button2.BackColor = colorDlg.Color;
                b2 = 1;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            Form1 fone = new Form1();
            Form2 f2 = new Form2();
            if(b1==1)
            {
                f1.tableLayoutPanel1.BackColor = Header;
                f2.tableLayoutPanel1.BackColor = Header;
                string HeaderHex=c.HexConverter(Header);

                dataStorage.setHEadercol(HeaderHex);

            }
           if(b2==1)
            {
                f1.tableLayoutPanel2.BackColor = Secondheader;

                f2.tableLayoutPanel2.BackColor = Secondheader;
                string SecondHeaderHex = c.HexConverter(Secondheader);
                

                dataStorage.setSecondcol(SecondHeaderHex);

            }
           if(richTextBox1.Text!="")
            {
                dataStorage.SloganText(richTextBox1.Text);
            }
            fone.Show();
          
            this.Close();
           

        }
    }
}
