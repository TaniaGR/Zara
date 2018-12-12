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

namespace Zara
{
    public partial class ychet_zal : Form
    {
        podklclss _CC;
        string id_tov_skld, mestop, id_tov_zal, post_id;
        int kol = 5;
        public ychet_zal()
        {
            InitializeComponent();
        }
        public void tov_sklad()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand tov_skld = new SqlCommand("select * from tov_skld_post_kl", _CC.conection);
            _CC.conection.Open();
            SqlDataReader tov = tov_skld.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(tov);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Артикул";
            dataGridView1.Columns[2].HeaderText = "Наименование товара";
            dataGridView1.Columns[3].HeaderText = "Количество";
            dataGridView1.Columns[4].HeaderText = "Цена";
            dataGridView1.Columns[5].HeaderText = "Место на складе";
            dataGridView1.Columns[6].Visible = false;
            dataGridView1.Columns[7].HeaderText = "Номер поставки";
            _CC.conection.Close();
        }

        public void tov_zal()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand tov_zal = new SqlCommand("select * from tov_zal_post", _CC.conection);
            _CC.conection.Open();
            SqlDataReader tov = tov_zal.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(tov);
            dataGridView2.DataSource = dt;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].HeaderText = "Артикул";
            dataGridView2.Columns[2].HeaderText = "Наименование товара";
            dataGridView2.Columns[3].HeaderText = "Количество";
            dataGridView2.Columns[4].HeaderText = "Цена";
            dataGridView2.Columns[5].Visible = false;
            dataGridView2.Columns[6].HeaderText = "Номер поставки";
            _CC.conection.Close();
        }
        
        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ychet_sklad examp = new ychet_sklad();
            examp.Show();
            this.Close();
        }

        private void вернутьсяВГлавноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }        

        private void главноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox10.Text != "")
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand insert_tov_zal = new SqlCommand("insert into [dbo].[tov_zal](artikyl,naim_tov_zal, kol_tov_zal, cena, post_zal_id) values(@artikyl,@naim_tov_zal,@kol_tov_zal,@cena,@post_zal_id)", _CC.conection);
                _CC.conection.Open();
                insert_tov_zal.Parameters.AddWithValue("artikyl", textBox10.Text);
                insert_tov_zal.Parameters.AddWithValue("naim_tov_zal", textBox1.Text);
                insert_tov_zal.Parameters.AddWithValue("kol_tov_zal", textBox2.Text);
                insert_tov_zal.Parameters.AddWithValue("cena", textBox5.Text);
                insert_tov_zal.Parameters.AddWithValue("post_zal_id", post_id);
                insert_tov_zal.ExecuteNonQuery();
                _CC.conection.Close();



                SqlConnection con = new SqlConnection("Data Source=DESKTOP-AG8AKLU;initial catalog=Zara_base;Persist Security info = True; User ID = SA; Password = qweqweqwe123");
                con.Open();
                SqlCommand delete_tov_sklad = new SqlCommand("[dbo].delete_tov_sklad", con);
                delete_tov_sklad.CommandType = CommandType.StoredProcedure;
                delete_tov_sklad.Parameters.AddWithValue("@id_tov_skld", Convert.ToInt32(id_tov_skld));
                delete_tov_sklad.ExecuteNonQuery();
                MessageBox.Show("Товар перемещен");
                textBox10.Text = "";
                textBox1.Text = "";
                textBox2.Text = "";
                textBox5.Text = "";
                textBox9.Text = "";
                tov_sklad();
                try
                {
                    _CC = new podklclss();
                    _CC.Set_Connection();
                    SqlCommand update_org_sklad = new SqlCommand("update org_sklad set [stat_id] = @stat_id where [id_mest_skl] = @id_mest_skl", _CC.conection);
                    _CC.conection.Open();
                    update_org_sklad.Parameters.AddWithValue("id_mest_skl", mestop);
                    update_org_sklad.Parameters.AddWithValue("stat_id", 1);
                    update_org_sklad.ExecuteNonQuery();
                    _CC.conection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                tov_sklad();
                tov_zal();
            }
            else
            {
                MessageBox.Show("Товар не выбран");
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox7.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }       

        private void ychet_zal_Load(object sender, EventArgs e)
        {
            tov_sklad();
            tov_zal();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id_tov_skld = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox10.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            //textBox2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox2.Text = Convert.ToString(kol);
            textBox5.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            mestop = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            post_id= dataGridView1.CurrentRow.Cells[6].Value.ToString();
            textBox9.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();            
        }
    }
}
