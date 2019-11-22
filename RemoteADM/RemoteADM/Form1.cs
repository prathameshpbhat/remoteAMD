using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
namespace RemoteADM
{
    public partial class Form1 : Form
    {
        ExcelFile c = new ExcelFile();
        int start=0;
        int l = 5;
        int cr = 0, kr = 0;
        String ET;
        struct database
        {
            public int state;
           public  String name;
        }
        database[,] array = new database[18, 6];
       
        Form2 dlg = new Form2();
        String name, room_number, bed_number, entry_time, departure_time, entry_date, Departure_date;

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (start != 0)
            {

                DataGridViewCellStyle style = new DataGridViewCellStyle();
                style.Font = new Font(dataGridView1.Font, FontStyle.Bold);

                style.ForeColor = Color.Black;
                dataGridView1.Rows[Int32.Parse(room_number) - 1].Cells[Int32.Parse(bed_number)].Style = style;
                dlg.dataGridView1.Rows[Int32.Parse(room_number) - 1].Cells[Int32.Parse(bed_number)].Style = style;

                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "OUT";
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Green;
                dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.ForeColor = Color.Black;
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Font = new Font("Segoe UI", 12, FontStyle.Bold);


                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Value = "Vacant";
                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.BackColor = Color.Green;
                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.ForeColor = Color.Black;
                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.Font = new Font("Segoe UI", 11, FontStyle.Bold);

                if (array[e.RowIndex, e.ColumnIndex].state == 1)
                {
                    cr--;
                    label_Cr.Text = cr.ToString();
                    dlg.label_Cr.Text = cr.ToString();
                    array[e.RowIndex, e.ColumnIndex].state = 0;


                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";

                    if (!System.IO.File.Exists(@"d:\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx"))
                    {

                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";
                        c.Rownumber = l;
                        c.create_newExcel();


                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "CR", "-----", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                    else
                    {
                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "CR", "--------", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                }
                else if (array[e.RowIndex, e.ColumnIndex].state == 2)
                {
                    kr--;
                    label_Kr.Text = kr.ToString();
                    dlg.label_Kr.Text = kr.ToString();
                    array[e.RowIndex, e.ColumnIndex].state = 0;

                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";

                    if (!System.IO.File.Exists(@"d:\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx"))
                    {

                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";
                        c.Rownumber = l;
                        c.create_newExcel();


                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "KR", "-----", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                    else
                    {
                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "KR", "--------", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
         
       

        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
          
            }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void timer3_Tick(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (start != 0)
            {

                DataGridViewCellStyle style = new DataGridViewCellStyle();
                style.Font = new Font(dataGridView1.Font, FontStyle.Bold);

                style.ForeColor = Color.Black;
                dataGridView1.Rows[Int32.Parse(room_number) - 1].Cells[Int32.Parse(bed_number)].Style = style;
                dlg.dataGridView1.Rows[Int32.Parse(room_number) - 1].Cells[Int32.Parse(bed_number)].Style = style;

                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "OUT";
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Green;
                dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.ForeColor = Color.Black;
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Font = new Font("Segoe UI", 12, FontStyle.Bold);


                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Value = "Vacant";
                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.BackColor = Color.Green;
                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.ForeColor = Color.Black;
                dlg.dataGridView1.Rows[e.ColumnIndex - 1].Cells[e.RowIndex + 1].Style.Font = new Font("Segoe UI", 11, FontStyle.Bold);

                if (array[e.RowIndex, e.ColumnIndex].state == 1)
                {
                    cr--;
                    label_Cr.Text = cr.ToString();
                    dlg.label_Cr.Text = cr.ToString();
                    array[e.RowIndex, e.ColumnIndex].state = 0;


                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";

                    if (!System.IO.File.Exists(@"d:\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx"))
                    {

                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";
                        c.Rownumber = l;
                        c.create_newExcel();


                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "CR", "-----", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                    else
                    {
                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "CR", "--------", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                }
                else if (array[e.RowIndex, e.ColumnIndex].state == 2)
                {
                    kr--;
                    label_Kr.Text = kr.ToString();
                    dlg.label_Kr.Text = kr.ToString();
                    array[e.RowIndex, e.ColumnIndex].state = 0;

                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";

                    if (!System.IO.File.Exists(@"d:\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx"))
                    {

                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";
                        c.Rownumber = l;
                        c.create_newExcel();


                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "KR", "-----", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                    else
                    {
                        c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.Rownumber = l;
                        c.openExcel();
                        String entry_time1 = DateTime.Now.ToString("HH:mm:ss tt").ToString();
                        String entry_date1 = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                        c.addDataToExcel(array[e.RowIndex, e.ColumnIndex].name, "KR", "--------", entry_time1 + entry_date1);
                        c.closeExcel();
                        l++;
                    }
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
          

           
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            for (int i = 0; i < 18; i++)
                for (int j = 0; j < 6; j++)
                {
                    array[i, j].state = 0;
                    array[i, j].name = "";
                }
                  

            label_Date.Text = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");


            label_Cr.Text = cr.ToString();

            label_Kr.Text = cr.ToString();

            dlg.label_Cr.Text = cr.ToString();

            dlg.label_Kr.Text = cr.ToString();

            try
            {
                Screen[] screens = Screen.AllScreens;
                Rectangle bounds = screens[1].Bounds;
                dlg.SetBounds(bounds.X, bounds.Y, bounds.Width, bounds.Height);
                dlg.StartPosition = FormStartPosition.Manual;

                dlg.Show();
            }
            catch
            {
                MessageBox.Show("Please connect monitor properly");

            }


            dlg.label_Date.Text = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");


            label_Time.Text = DateTime.Now.ToString("HH:mm:ss tt").ToString();
            dlg.label_Time.Text = DateTime.Now.ToString("HH:mm:ss tt").ToString();


           

          
            var time = DateTime.Now;
            string formattedTime = time.ToString(" hh, mm, ss,dd,MM,yyyy");
           
            dlg.dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToAddRows = true;
            for (int i=0;i<18;i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = i + 1;
                for (int j=1;j<=6;j++)
                {
              
                  
                    dataGridView1.Rows[i].Cells[j].Value = "OUT";
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Green;
                    dataGridView1.Rows[i].Cells[j].Style.Font = new Font("Segoe UI", 12, FontStyle.Bold);

                }
                 
            }
            dataGridView1.Rows[0].Cells[0].Value = 1;


            for (int i = 0; i < 6; i++)
            {
              
                dlg.dataGridView1.Rows.Add();
                dlg.dataGridView1.Rows[i].Cells[0].Style.Font = new Font("Segoe UI", 15, FontStyle.Bold);
                dlg.dataGridView1.Rows[i].Cells[0].Value ="Bed"+ (i + 1);
                for (int j = 1; j <= 18; j++)
                {
                   
                
                    dlg.dataGridView1.Rows[i].Cells[j].Value = "Vacant";
                    dlg.dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Green;
                   
                    dlg.dataGridView1.Rows[i].Cells[j].Style.Font = new Font("Segoe UI", 11, FontStyle.Bold);

                }

            }
            dlg.dataGridView1.Rows[0].Cells[0].Value = "Bed" + 1;
        }
       
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        void setdata()
        {
            dataGridView1.Rows[Int32.Parse(room_number) - 1].Cells[Int32.Parse(bed_number)].Value = name;
            DataGridViewCellStyle style = new DataGridViewCellStyle();
            style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
            style.BackColor = Color.Orange; 
            style.ForeColor = Color.White;
            dataGridView1.Rows[Int32.Parse(room_number) - 1].Cells[Int32.Parse(bed_number)].Style = style;


            dlg.dataGridView1.Rows[Int32.Parse(bed_number) - 1].Cells[Int32.Parse(room_number)].Value = name;
            DataGridViewCellStyle style1 = new DataGridViewCellStyle();
            style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
            style.BackColor = Color.FromArgb(255, 77, 77);
            style.ForeColor = Color.White;
            dlg.dataGridView1.Rows[Int32.Parse(bed_number) - 1].Cells[Int32.Parse(room_number)].Style = style;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            start = 0;

            name = Text_name.Text;
            bed_number = Text_Bed_NO.Text;
            room_number = Text_Room_No.Text;
            entry_time = DateTime.Now.ToString("HH:mm:ss tt").ToString();
            entry_date = DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
      
                if(radio_Cr.Checked==true&& radio_Kr.Checked == true)
            {
                MessageBox.Show("You  Can't Select both CR and KR");
            }
                else if(radio_Cr.Checked == true)
            {
               
                array[Int32.Parse(room_number)-1, Int32.Parse(bed_number)].name = name;
                array[Int32.Parse(room_number)-1, Int32.Parse(bed_number)].state = 1;
                cr++;

                label_Cr.Text = cr.ToString();

                dlg.label_Cr.Text = cr.ToString();
                setdata();
                c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy")+".xlsx";

                if (!System.IO.File.Exists(@"d:\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx"))
                {

                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";
                    c.Rownumber = l;
                    c.create_newExcel();
                    

                    c.ExcelFilePath = "d:\\"+ DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                    c.Rownumber = l;
                    c.openExcel();
                    c.addDataToExcel(name,"CR", entry_time+ entry_date,"-------");
                    c.closeExcel();
                    l++;
                }
                else
                {
                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                    c.Rownumber = l;
                    c.openExcel();
                    c.addDataToExcel(name, "CR", entry_time + entry_date, "-------");
                    c.closeExcel();
                    l++;
                }
                
            }
            else if (radio_Kr.Checked == true)
            {
                array[Int32.Parse(room_number) - 1, Int32.Parse(bed_number) ].name = name;
                array[Int32.Parse(room_number)-1, Int32.Parse(bed_number)].state = 2;
                kr++;
                setdata();
                label_Kr.Text = kr.ToString();

                dlg.label_Kr.Text = kr.ToString();


                setdata();
                c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";

                if (!System.IO.File.Exists(@"d:\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx"))
                {

                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy") + ".xlsx";
                    c.Rownumber = l;
                    c.create_newExcel();


                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                    c.Rownumber = l;
                    c.openExcel();
                    c.addDataToExcel(name, "KR", entry_time + entry_date, "-------");
                    c.closeExcel();
                    l++;
                }
                else
                {
                    c.ExcelFilePath = "d:\\" + DateTime.UtcNow.Date.ToString("dd MMMM yyyy");
                    c.Rownumber = l;
                    c.openExcel();
                    c.addDataToExcel(name, "KR", entry_time + entry_date, "-------");
                    c.closeExcel();
                    l++;
                }
            }




            start = 1;


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label_Time.Text = DateTime.Now.ToString("HH:mm:ss tt").ToString();
            dlg.label_Time.Text = DateTime.Now.ToString("HH:mm:ss tt").ToString();
        }
    }
}
