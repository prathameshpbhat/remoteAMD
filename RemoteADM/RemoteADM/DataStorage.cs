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
using System.Data.SqlClient;
namespace RemoteADM
{
    class DataStorage
    {
        public void putdata(int room_id,int state,string name="",string Entry_time="", string Exit_time="")
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");

            con.Open();
            SqlCommand cmd = new SqlCommand("INSERT INTO [Table](room_id,state,Name,Entry_time,Exit_time) VALUES(@room_id,@state,@Name,@Entry_time,@Exit_time)", con);
      
            cmd.Parameters.AddWithValue("@room_id", room_id);
            cmd.Parameters.AddWithValue("@state", state);
            cmd.Parameters.AddWithValue("@Name", name);
            cmd.Parameters.AddWithValue("@Entry_time", Entry_time);

            cmd.Parameters.AddWithValue("@Exit_time", Exit_time);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void getdata()
        {

        }
        public void updateData()
        {

        }
        public void getentiredata(Form2 dlg, Form1 datagrid1)
        {
        
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand("select * from [Table] ", con);

            SqlDataReader sdr = command.ExecuteReader();
         
                while (sdr.Read())
                {
                    
                  
                    int val = Int32.Parse(sdr[0].ToString());
                    int bed_num = val % 10;
                    val = val / 10;
                    int room_number = val % 10;
                    datagrid1.dataGridView1.Rows[room_number].Cells[bed_num+1].Value = sdr["Name"].ToString(); 
                    DataGridViewCellStyle style = new DataGridViewCellStyle();
                    style.Font = new Font(datagrid1.dataGridView1.Font, FontStyle.Bold);
                    style.BackColor = Color.Orange;
                    style.ForeColor = Color.White;
                    datagrid1.dataGridView1.Rows[room_number ].Cells[bed_num+1].Style = style;
                    Form1.array[room_number, bed_num].state= Int32.Parse(sdr["state"].ToString());
                    Form1.array[room_number, bed_num].name = sdr["Name"].ToString();
                    
                    dlg.dataGridView1.Rows[bed_num ].Cells[room_number+1].Value = sdr["Name"].ToString(); ;
                    DataGridViewCellStyle style1 = new DataGridViewCellStyle();
                    style.Font = new Font(dlg.dataGridView1.Font, FontStyle.Bold);
                    style.BackColor = Color.FromArgb(255, 77, 77);
                    style.ForeColor = Color.White;
                    dlg.dataGridView1.Rows[bed_num].Cells[room_number+1].Style = style;
                }

                con.Close();
            
           

            con.Close();
        }
       public void connect()
        {

            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");


            con.Open();
            MessageBox.Show("Connected successfully");
            con.Close();
        }
        public void Delete(int val)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
           // int.Parse(paramVal)
            //SqlCommand cmd = new SqlCommand("DELETE  from [Table] WHERE room_id = '@val' ", con);
            SqlCommand cmd = new SqlCommand("DELETE  from [Table] WHERE room_id = " + val+";", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void setCRKR(Form2 dlg, Form1 datagrid1)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
            SqlCommand command1 = new SqlCommand("select count(state) from [Table] where state=1 GROUP BY state", con);

            SqlDataReader sdr1 = command1.ExecuteReader();
         
            if (sdr1.Read())
            {
              

                Form1.cr = Int32.Parse(sdr1[0].ToString());
                
                con.Close();


            }
            else
            {
                con.Close();

            }
            con.Open();
            SqlCommand command2 = new SqlCommand("select count(state) from [Table] where state=2 GROUP BY state", con);
      
            SqlDataReader sdr2 = command2.ExecuteReader();
          
             if(sdr2.Read())
            {
              

                   Form1.kr = Int32.Parse(sdr2[0].ToString());

                con.Close();


            }
            

            con.Close();
        }
       public void setHEadercol(string headcol)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");

            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE   [Details] SET headcol=@headcol WHERE id=1", con);

            cmd.Parameters.AddWithValue("@headcol", headcol);
         
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void setSecondcol( string secondheadcol)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");

            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE   [Details] SET secondheadcol=@secondheadcol WHERE id=1", con);

          
            cmd.Parameters.AddWithValue("@secondheadcol", secondheadcol);
           

            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void SloganText(String slogan)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");

            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE   [Details] SET slogan=@slogan WHERE id=1", con);

            cmd.Parameters.AddWithValue("@slogan", slogan);

            cmd.ExecuteNonQuery();
            con.Close();
        }
        public string getheadcol()
        {

            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand("SELECT * FROM  [Details]  ", con);

            SqlDataReader sdr = command.ExecuteReader();
            if (sdr.Read())
                return sdr["headcol"].ToString();
            else
                return "#1A75FF";
        }
        public string getsecondcol()
        {

            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand("select * from [Details] where Id=1 ", con);

            SqlDataReader sdr = command.ExecuteReader();
            if (sdr.Read())
                return sdr["secondheadcol"].ToString();
            else
                return "#4D90FF";
           
        }
        public string sloganString()
        {

            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand("select * from [Details] where Id=1 ", con);

            SqlDataReader sdr = command.ExecuteReader();
            if (sdr.Read())
                return sdr["slogan"].ToString();
            else
                return "Don't be safely blinded be safety minded.";
        }

        public int getRownum()
        {

            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");
            con.Open();
            SqlCommand command = new SqlCommand("select * from [Details] where Id=1 ", con);

            SqlDataReader sdr = command.ExecuteReader();
            if (sdr.Read())
                return Int32.Parse(sdr["Rownum"].ToString());
     
            else
                return 1;
        }
        public void SetRownum(int Rownum)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Lenovo\documents\visual studio 2017\Projects\RemoteADM\RemoteADM\Database1.mdf;Integrated Security=True");

            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE   [Details] SET Rownum=@Rownum WHERE id=1", con);

            cmd.Parameters.AddWithValue("@Rownum", Rownum);

            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}
