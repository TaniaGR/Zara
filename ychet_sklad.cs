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
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Zara
{
    public partial class ychet_sklad : Form
    {
        podklclss _CC;
        int nom_pst;
        int kol = 5;
        string id_post, id_tov_skld, mesto, mestop;
        private readonly string TenplateFileName = @"C:\Users\house\Desktop\Образование\РВиАПООН Зеленина И-1-15\КП 02 01. Зеленина И-1-15\Программа\Zara\Шаблоны\Акт поставки.dotx";
        private readonly string TenplateFileName2 = @"C:\Users\house\Desktop\Образование\РВиАПООН Зеленина И-1-15\КП 02 01. Зеленина И-1-15\Программа\Zara\Шаблоны\Акт списания.dotx";
        public ychet_sklad()
        {
            InitializeComponent();
        }

        public void post()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand post = new SqlCommand("select * from post", _CC.conection);
            _CC.conection.Open();
            SqlDataReader pst = post.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(pst);
            dataGridView3.DataSource = dt;
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[1].HeaderText = "Номер поставки";
            dataGridView3.Columns[2].HeaderText = "Дата поставки";
            _CC.conection.Close();
        }

        public void get_nom_post() 
        {            
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand cmd = new SqlCommand("select count([nom_post])+1 from post", _CC.conection);
            _CC.conection.Open();
            maskedTextBox3.Text = cmd.ExecuteScalar().ToString();
            nom_pst = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            _CC.conection.Close();            
        }

        public void mest_skld()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand mest_skld = new SqlCommand("select * from mst_skld", _CC.conection);
            _CC.conection.Open();
            SqlDataReader mest = mest_skld.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(mest);
            dataGridView2.DataSource = dt;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].HeaderText = "Номер места";
            dataGridView2.Columns[2].HeaderText = "Номер стелажа";
            dataGridView2.Columns[3].HeaderText = "Номер по верт";
            dataGridView2.Columns[4].HeaderText = "Номер по гор";
            dataGridView2.Columns[5].HeaderText = "Статус";
            _CC.conection.Close();
        }        

        public void tov_sklad()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand tov_skld = new SqlCommand("select * from tov_skld_post", _CC.conection);
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
            dataGridView1.Columns[6].HeaderText = "Номер поставки";
            _CC.conection.Close();
        }

        private void торговыйЗалToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ychet_zal examp = new ychet_zal();
            examp.Show();
            this.Close();
        }

        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void вернутьсяВГлавноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {            
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand update_tov_sklad = new SqlCommand("update tov_sklad set [artikyl] = @artikyl, [naim_tov_skld] = @naim_tov_skld, [kol_tov_skld]=@kol_tov_skld, [cena]=@cena, [mest_skl_id]=@mest_skl_id, [post_id]=@post_id where [id_tov_skld] = @id_tov_skld", _CC.conection);
                _CC.conection.Open();
                update_tov_sklad.Parameters.AddWithValue("id_tov_skld", id_tov_skld);
                update_tov_sklad.Parameters.AddWithValue("artikyl", textBox5.Text);
                update_tov_sklad.Parameters.AddWithValue("naim_tov_skld",textBox1.Text);
                update_tov_sklad.Parameters.AddWithValue("kol_tov_skld", textBox2.Text);
                update_tov_sklad.Parameters.AddWithValue("cena", textBox4.Text);
                update_tov_sklad.Parameters.AddWithValue("mest_skl_id", comboBox2.SelectedValue);
                update_tov_sklad.Parameters.AddWithValue("post_id", comboBox1.SelectedValue);
                update_tov_sklad.ExecuteNonQuery();
                _CC.conection.Close();
                tov_sklad();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if (mesto != comboBox2.Text)
            {
                try
                {
                    int m = 2;
                    _CC = new podklclss();
                    _CC.Set_Connection();
                    SqlCommand update_org_sklad = new SqlCommand("update org_sklad set [stat_id] = @stat_id where [id_mest_skl] = @id_mest_skl", _CC.conection);
                    _CC.conection.Open();
                    update_org_sklad.Parameters.AddWithValue("id_mest_skl", comboBox2.SelectedValue);
                    update_org_sklad.Parameters.AddWithValue("stat_id", m);
                    update_org_sklad.ExecuteNonQuery();
                    _CC.conection.Close();
                    mest_skld();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                try
                {
                    int m = 1;
                    _CC = new podklclss();
                    _CC.Set_Connection();
                    SqlCommand update_org_sklad = new SqlCommand("update org_sklad set [stat_id] = @stat_id where [id_mest_skl] = @id_mest_skl", _CC.conection);
                    _CC.conection.Open();
                    update_org_sklad.Parameters.AddWithValue("id_mest_skl", mestop);
                    update_org_sklad.Parameters.AddWithValue("stat_id", 1);
                    update_org_sklad.ExecuteNonQuery();
                    _CC.conection.Close();
                    mest_skld();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }       
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void com_post()
        {            
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_base2DataSet3.post". При необходимости она может быть перемещена или удалена.
            this.postTableAdapter.Fill(this.zara_base2DataSet3.post);
        }

        private void ychet_sklad_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_baseDataSet2.org_sklad". При необходимости она может быть перемещена или удалена.
            this.org_skladTableAdapter2.Fill(this.zara_baseDataSet2.org_sklad);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_baseDataSet1.post". При необходимости она может быть перемещена или удалена.
            this.postTableAdapter1.Fill(this.zara_baseDataSet1.post);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_base2DataSet5.org_sklad". При необходимости она может быть перемещена или удалена.
            this.org_skladTableAdapter1.Fill(this.zara_base2DataSet5.org_sklad);
            com_post();
            // рныTODO: данная строка кода позволяет загрузить данные в таблицу "zara_base2DataSet3.post". При необходимости она может быть перемещена или удалена.
            //this.postTableAdapter.Fill(this.zara_base2DataSet3.post);
            post();
            get_nom_post();
            mest_skld();
            tov_sklad();
        }        

        private void главноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand insert_post = new SqlCommand("insert into [dbo].[post](nom_post,data_post) values(@nom_post, @data_post)", _CC.conection);
                _CC.conection.Open();
                insert_post.Parameters.AddWithValue("nom_post", nom_pst);
                insert_post.Parameters.AddWithValue("data_post", maskedTextBox4.Text);
                insert_post.ExecuteNonQuery();
                _CC.conection.Close();
                maskedTextBox4.Text = "";
                get_nom_post();
                post();
                com_post();
            }
            catch
            {
                MessageBox.Show("Заполните дату поставки");
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id_post = dataGridView3.CurrentRow.Cells[0].Value.ToString();
            maskedTextBox3.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();            
            maskedTextBox4.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();            
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++) //row-строка
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox3.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection("Data Source=DESKTOP-AG8AKLU;initial catalog=Zara_base;Persist Security info = True; User ID = SA; Password = qweqweqwe123");
                con.Open();
                SqlCommand delete_tov_sklad = new SqlCommand("[dbo].delete_tov_sklad", con);
                delete_tov_sklad.CommandType = CommandType.StoredProcedure;
                delete_tov_sklad.Parameters.AddWithValue("@id_tov_skld", Convert.ToInt32(id_tov_skld));
                delete_tov_sklad.ExecuteNonQuery();
                MessageBox.Show("Товар удален");
                textBox5.Text = "";
                textBox1.Text = "";
                textBox4.Text = "";
                comboBox1.Text = "";
                comboBox2.Text = "";
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
                    mest_skld();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch
            {
                MessageBox.Show("Выберите удаляемый товар");
            }
        }

        private void maskedTextBox4_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
            {
                e.Handled = true;
            }
        }

        private void ReplaceWordStub(string stud, string text, Word.Document wordDocument)
        {
            //string data = Convert.ToString(DateTime.Today);
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stud, ReplaceWith: text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (maskedTextBox4.Text != "" && maskedTextBox3.Text != "") 
            {
                string data;
                DateTime dt = DateTime.Now;
                data = dt.ToShortDateString().ToString();
                var wordApp = new Word.Application();
                wordApp.Visible = false;
                try
                {
                    var wordDociment = wordApp.Documents.Open(TenplateFileName);
                    ReplaceWordStub("@data", data, wordDociment);
                    ReplaceWordStub("@nomer", maskedTextBox3.Text, wordDociment);
                    ReplaceWordStub("@data", data, wordDociment);
                    ReplaceWordStub("@nomer", maskedTextBox3.Text, wordDociment);                 

                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите поставку");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox5.Text != "" && textBox1.Text != "" && textBox4.Text != "")
            {
                string data;
                DateTime dt = DateTime.Now;
                data = dt.ToShortDateString().ToString();
                var wordApp = new Word.Application();
                wordApp.Visible = false;
                try
                {
                    var wordDociment = wordApp.Documents.Open(TenplateFileName2);
                    ReplaceWordStub("@data", data, wordDociment);
                    ReplaceWordStub("@data", data, wordDociment);
                    ReplaceWordStub("@art", textBox5.Text, wordDociment);
                    ReplaceWordStub("@naim", textBox1.Text, wordDociment);
                    ReplaceWordStub("@cena", textBox4.Text, wordDociment);
                    
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите поставку");
            }
        }

        private void ychet_sklad_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand update_post = new SqlCommand("update post set [nom_post] = @nom_post,[data_post] = @data_post where [id_post] = @id_post", _CC.conection);
                _CC.conection.Open();
                update_post.Parameters.AddWithValue("id_post", id_post);
                update_post.Parameters.AddWithValue("nom_post", maskedTextBox3.Text);
                update_post.Parameters.AddWithValue("data_post", maskedTextBox4.Text);
                update_post.ExecuteNonQuery();
                _CC.conection.Close();
                post();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox5.Text != "" && textBox1.Text != "" && textBox2.Text != "" && textBox4.Text != "")
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand insert_tov_skld = new SqlCommand("insert into [dbo].[tov_sklad](artikyl,naim_tov_skld, kol_tov_skld, cena, mest_skl_id, post_id) values(@artikyl,@naim_tov_skld,@kol_tov_skld,@cena,@mest_skl_id,@post_id)", _CC.conection);
                _CC.conection.Open();
                insert_tov_skld.Parameters.AddWithValue("artikyl", textBox5.Text);
                insert_tov_skld.Parameters.AddWithValue("naim_tov_skld", textBox1.Text);
                insert_tov_skld.Parameters.AddWithValue("kol_tov_skld", textBox2.Text);
                insert_tov_skld.Parameters.AddWithValue("cena", textBox4.Text);
                insert_tov_skld.Parameters.AddWithValue("mest_skl_id", comboBox2.SelectedValue);
                insert_tov_skld.Parameters.AddWithValue("post_id", comboBox1.SelectedValue);
                insert_tov_skld.ExecuteNonQuery();
                _CC.conection.Close();
                tov_sklad();

                try
                {
                    int m = 2;
                    _CC = new podklclss();
                    _CC.Set_Connection();
                    SqlCommand update_org_sklad = new SqlCommand("update org_sklad set [stat_id] = @stat_id where [id_mest_skl] = @id_mest_skl", _CC.conection);
                    _CC.conection.Open();
                    update_org_sklad.Parameters.AddWithValue("id_mest_skl", comboBox2.SelectedValue);
                    update_org_sklad.Parameters.AddWithValue("stat_id", m);
                    update_org_sklad.ExecuteNonQuery();
                    _CC.conection.Close();
                    mest_skld();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Заполните данные о товаре");
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id_tov_skld = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            //textBox2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox2.Text = Convert.ToString(kol);
            textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            mesto = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            mestop = Convert.ToString(comboBox2.SelectedValue);

        }
    }
}
