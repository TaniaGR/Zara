using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace Zara
{
    public partial class kadr_rab_time : Form
    {
        string id_dolzh, id_otpyska, id_bol;
        int otp, rab, otrab;        
        podklclss _CC;
         
        public kadr_rab_time()
        {
            InitializeComponent();
        }

        public void load_dolzhn()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand load_dolzh = new SqlCommand("select * from [dbo].[dolzhn]", _CC.conection);
            _CC.conection.Open();
            SqlDataReader dr = load_dolzh.ExecuteReader();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Должность";
            dataGridView1.Columns[2].HeaderText = "Оклад";
            dataGridView1.Columns[3].HeaderText = "Рабочие часы";
            _CC.conection.Close();
        }
       
        public void load_otpysk()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand load_otpysk = new SqlCommand("select * from [dbo].[sotr_otp]", _CC.conection);
            _CC.conection.Open();
            SqlDataReader dr = load_otpysk.ExecuteReader();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);
            dataGridView2.DataSource = dt;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].HeaderText = "Табельный номер";
            dataGridView2.Columns[2].HeaderText = "Дата с";
            dataGridView2.Columns[3].HeaderText = "Дата по";
            _CC.conection.Close();
        }

        private void больничныеЛистыToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            kadr_sotr examp = new kadr_sotr();
            examp.Show();
            this.Close();
        }

        private void kadr_rab_time_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_base2DataSet.sotrud". При необходимости она может быть перемещена или удалена.
            this.sotrudTableAdapter.Fill(this.zara_base2DataSet.sotrud);
            load_dolzhn();
            load_otpysk();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id_dolzh = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text!="" && textBox3.Text!="")
            {
                try
                {
                    _CC = new podklclss();
                    _CC.Set_Connection();
                    SqlCommand insert_sotr = new SqlCommand("insert into [dbo].[dolzhn](naim_dolzh, oklad, rab_chas) " +
                        "values(@naim_dolzh, @oklad, @rab_chas)", _CC.conection);
                    _CC.conection.Open();
                    insert_sotr.Parameters.AddWithValue("naim_dolzh", textBox1.Text);
                    insert_sotr.Parameters.AddWithValue("oklad", textBox2.Text);
                    insert_sotr.Parameters.AddWithValue("rab_chas", textBox3.Text);
                    insert_sotr.ExecuteNonQuery();
                    _CC.conection.Close();
                    load_dolzhn();
                }
                catch
                {
                    MessageBox.Show("Ошибка добавления");
                }
            }
            else
            {
                MessageBox.Show("Заполните все поля");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand update_dolzh = new SqlCommand("update [dbo].[dolzhn] set [naim_dolzh] = @naim_dolzh,[oklad] = @oklad," +
                    "[rab_chas] = @rab_chas where [id_dolzh] = @id_dolzh", _CC.conection);
                _CC.conection.Open();
                update_dolzh.Parameters.AddWithValue("id_dolzh", id_dolzh);
                update_dolzh.Parameters.AddWithValue("naim_dolzh", textBox1.Text);
                update_dolzh.Parameters.AddWithValue("oklad", textBox2.Text);
                update_dolzh.Parameters.AddWithValue("rab_chas", textBox3.Text);
                update_dolzh.ExecuteNonQuery();
                _CC.conection.Close();
                load_dolzhn();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=DESKTOP-AG8AKLU;initial catalog=Zara_base2;Persist Security info = True; User ID = SA; Password = qweqweqwe123");
                con.Open();
                SqlCommand delete_dolzh = new SqlCommand("[dbo].delete_dolzhn", con);
                delete_dolzh.CommandType = CommandType.StoredProcedure;
                delete_dolzh.Parameters.AddWithValue("@id_dolzh", Convert.ToInt32(id_dolzh));
                delete_dolzh.ExecuteNonQuery();
                MessageBox.Show("Должность удалена");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                load_dolzhn();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Данную должность нельзя удалить "+ ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand insert_otp = new SqlCommand("insert into [dbo].[otpysk](data_s, data_po, sotr_otp_id) " +
                    "values(@data_s, @data_po, @sotr_otp_id)", _CC.conection);
                _CC.conection.Open();
                insert_otp.Parameters.AddWithValue("data_s", maskedTextBox1.Text);
                insert_otp.Parameters.AddWithValue("data_po", maskedTextBox2.Text);
                insert_otp.Parameters.AddWithValue("sotr_otp_id", comboBox1.SelectedValue);
                insert_otp.ExecuteNonQuery();
                _CC.conection.Close();
                load_otpysk();
            }
            catch
            {
                MessageBox.Show("Заполните все поля");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand update_otp = new SqlCommand("update [dbo].[otpysk] set [data_s] = @data_s,[data_po] = @data_po," +
                    "[sotr_otp_id] = @sotr_otp_id, [id_otpyska] = @id_otpyska", _CC.conection);
                _CC.conection.Open();
                update_otp.Parameters.AddWithValue("id_otpyska", id_otpyska);
                update_otp.Parameters.AddWithValue("data_s", maskedTextBox1.Text);
                update_otp.Parameters.AddWithValue("data_po", maskedTextBox2.Text);
                update_otp.Parameters.AddWithValue("sotr_otp_id", comboBox1.SelectedValue);
                update_otp.ExecuteNonQuery();
                _CC.conection.Close();
                load_dolzhn();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=DESKTOP-AG8AKLU;initial catalog=Zara_base2;Persist Security info = True; User ID = SA; Password = qweqweqwe123");
                con.Open();
                SqlCommand delete_otpysk = new SqlCommand("[dbo].delete_otpysk", con);
                delete_otpysk.CommandType = CommandType.StoredProcedure;
                delete_otpysk.Parameters.AddWithValue("@id_otpyska", Convert.ToInt32(id_otpyska));
                delete_otpysk.ExecuteNonQuery();
                MessageBox.Show("Данные об отпуске удалены");
                comboBox1.Text = "";
                maskedTextBox1.Text = "";
                maskedTextBox2.Text = "";
                load_otpysk();
            }
            catch
            {
                MessageBox.Show("Выберите данные отпуска");
            }
        }                  
        
        private void button5_Click_1(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand update_dolzh = new SqlCommand("update [dbo].[otpysk] set [data_s] = @data_s,[data_po] = @data_po," +
                    "[sotr_otp_id] = @sotr_otp_id where [id_otpyska] = @id_otpyska", _CC.conection);
                _CC.conection.Open();
                update_dolzh.Parameters.AddWithValue("id_otpyska", id_otpyska);
                update_dolzh.Parameters.AddWithValue("data_s", maskedTextBox1.Text);
                update_dolzh.Parameters.AddWithValue("data_po", maskedTextBox2.Text);
                update_dolzh.Parameters.AddWithValue("sotr_otp_id", comboBox1.SelectedValue);
                update_dolzh.ExecuteNonQuery();
                _CC.conection.Close();
                load_otpysk();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand otp_c = new SqlCommand("select datediff(day,data_s,data_po) from otpysk where id_otpyska='" + id_otpyska + "'", _CC.conection);
                _CC.conection.Open();
                otp = Convert.ToInt32(otp_c.ExecuteScalar().ToString());
                _CC.conection.Close();
                rab = Convert.ToInt32(textBox3.Text);
        
                otrab = (rab - otp);
                if (otrab < 0)
                {
                    if (otrab > (-20))
                    {
                        rab = 40;
                        otrab = (rab - otp);
                        MessageBox.Show("Результаты показаны за 2 месяца");
                    }
                }
                textBox6.Text = otp.ToString();
                textBox5.Text = otrab.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Выберите должность и данные об отпуске "+ex.Message);
            }
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void главноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void maskedTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
           Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application(); 
            ExcelApp.Application.Workbooks.Add(Type.Missing); 
            ExcelApp.Columns.ColumnWidth = 25;            
             
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[j].Cells[i].Value;
                }
            }
           ExcelApp.Visible = true;
           ExcelApp.UserControl = true;
        }

        private void dataGridView2_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            id_otpyska = dataGridView2.CurrentRow.Cells[0].Value.ToString();
            comboBox1.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
            maskedTextBox1.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
            maskedTextBox2.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
        }

        private void kadr_rab_time_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.RowCount; i++) 
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox7.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }       

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }    
                
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
            
        }
    }
}
