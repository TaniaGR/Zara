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
    public partial class kadr_sotr : Form
    {
        private readonly string TenplateFileName = @"C:\Users\house\Desktop\Образование\РВиАПООН Зеленина И-1-15\КП 02 01. Зеленина И-1-15\Программа\Zara\Шаблоны\Трудовой договор.dotx";
        private readonly string TenplateFileName2 = @"C:\Users\house\Desktop\Образование\РВиАПООН Зеленина И-1-15\КП 02 01. Зеленина И-1-15\Программа\Zara\Шаблоны\Приказ на увольнение.dotx";
        string id_sotr;
        int tabel_num;
        Word._Application oWord = new Word.Application();
        object oMissing = System.Reflection.Missing.Value;
        podklclss _CC;      

        public kadr_sotr()
        {
            InitializeComponent();
        }       

        public void load_sotr()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand load_sotr = new SqlCommand("select * from [dbo].[sotr_dolzh]", _CC.conection);
            _CC.conection.Open();
            SqlDataReader dr = load_sotr.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(dr);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Табельный";
            dataGridView1.Columns[2].HeaderText = "Фамилия";
            dataGridView1.Columns[3].HeaderText = "Имя";
            dataGridView1.Columns[4].HeaderText = "Отчетсво";
            dataGridView1.Columns[5].HeaderText = "Дата рождения";
            dataGridView1.Columns[6].HeaderText = "Серия паспорта";
            dataGridView1.Columns[7].HeaderText = "Номер паспорта";
            dataGridView1.Columns[8].HeaderText = "Номер трудовой книжки";
            dataGridView1.Columns[9].HeaderText = "Дата получения трудовой";
            dataGridView1.Columns[10].HeaderText = "Логин";
            dataGridView1.Columns[11].HeaderText = "Пароль";
            dataGridView1.Columns[12].HeaderText = "СНИЛС";
            dataGridView1.Columns[13].HeaderText = "Кодовое слово";
            dataGridView1.Columns[14].HeaderText = "Должность";
            _CC.conection.Close();
        }

        private void больничныеЛистыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            kadr_rab_time examp = new kadr_rab_time();
            examp.Show();
            this.Close();            
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }

        private void kadr_sotr_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_baseDataSet.dolzhn". При необходимости она может быть перемещена или удалена.
            this.dolzhnTableAdapter1.Fill(this.zara_baseDataSet.dolzhn);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "zara_base2DataSet2.dolzhn". При необходимости она может быть перемещена или удалена.
            this.dolzhnTableAdapter.Fill(this.zara_base2DataSet2.dolzhn);
            load_sotr();
            get_tabel();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id_sotr = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            maskedTextBox9.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            maskedTextBox7.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            maskedTextBox8.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            maskedTextBox5.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            maskedTextBox2.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            maskedTextBox4.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            maskedTextBox6.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            maskedTextBox3.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand insert_sotr = new SqlCommand("insert into [dbo].[sotrud](tabel,familia,imya,otchestvo,data_rozh,seria_pas,nomer_pas,nomer_TK,data_pr_TK,login,parol,slovo,snils,dolzh_id) " +
                    "values(@tabel,@familia,@imya,@otchestvo,@data_rozh,@seria_pas,@nomer_pas,@nomer_TK,@data_pr_TK,@login,@parol,@slovo,@snils,@dolzh_id)", _CC.conection);
                _CC.conection.Open();
                insert_sotr.Parameters.AddWithValue("tabel", tabel_num);
                insert_sotr.Parameters.AddWithValue("familia", textBox1.Text);
                insert_sotr.Parameters.AddWithValue("imya", textBox2.Text);
                insert_sotr.Parameters.AddWithValue("otchestvo", textBox3.Text);
                insert_sotr.Parameters.AddWithValue("data_rozh", Convert.ToDateTime(maskedTextBox1.Text));
                insert_sotr.Parameters.AddWithValue("seria_pas", maskedTextBox7.Text);
                insert_sotr.Parameters.AddWithValue("nomer_pas", maskedTextBox8.Text);
                insert_sotr.Parameters.AddWithValue("nomer_TK", maskedTextBox5.Text);
                insert_sotr.Parameters.AddWithValue("data_pr_TK", maskedTextBox2.Text);
                insert_sotr.Parameters.AddWithValue("login", maskedTextBox4.Text);
                insert_sotr.Parameters.AddWithValue("parol", maskedTextBox6.Text);
                insert_sotr.Parameters.AddWithValue("slovo", textBox4.Text);
                insert_sotr.Parameters.AddWithValue("snils", maskedTextBox3.Text);
                insert_sotr.Parameters.AddWithValue("dolzh_id", comboBox1.SelectedValue);
                insert_sotr.ExecuteNonQuery();
                _CC.conection.Close();
                load_sotr();
            }
            catch
            {
                MessageBox.Show("Заполните все поля");
            }
        }

        public void get_tabel() 
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand cmd = new SqlCommand("select count([tabel])+1 from [dbo].[sotrud]", _CC.conection);
            _CC.conection.Open();
            maskedTextBox9.Text = cmd.ExecuteScalar().ToString();
            tabel_num = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            _CC.conection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                SqlCommand update_sotr = new SqlCommand("update [dbo].[sotrud] set [familia] = @familia,[imya] = @imya,[otchestvo] = @otchestvo,[data_rozh] = @data_rozh, " +
                    "[seria_pas] = @seria_pas, [nomer_pas] = @nomer_pas, [nomer_TK] = @nomer_TK,[data_pr_TK]=@data_pr_TK,[login]=@login,[parol]=@parol,[slovo]=@slovo, " +
                    "[snils]=@snils,[dolzh_id] = @dolzh_id where [id_sotr] = @id_sotr", _CC.conection);
                _CC.conection.Open();
                update_sotr.Parameters.AddWithValue("id_sotr", id_sotr);
                update_sotr.Parameters.AddWithValue("tabel", tabel_num);
                update_sotr.Parameters.AddWithValue("familia", textBox1.Text);
                update_sotr.Parameters.AddWithValue("imya", textBox2.Text);
                update_sotr.Parameters.AddWithValue("otchestvo", textBox3.Text);
                update_sotr.Parameters.AddWithValue("data_rozh", Convert.ToDateTime(maskedTextBox1.Text));
                update_sotr.Parameters.AddWithValue("seria_pas", maskedTextBox7.Text);
                update_sotr.Parameters.AddWithValue("nomer_pas", maskedTextBox8.Text);
                update_sotr.Parameters.AddWithValue("nomer_TK", maskedTextBox5.Text);
                update_sotr.Parameters.AddWithValue("data_pr_TK", maskedTextBox2.Text);
                update_sotr.Parameters.AddWithValue("login", maskedTextBox4.Text);
                update_sotr.Parameters.AddWithValue("parol", maskedTextBox6.Text);
                update_sotr.Parameters.AddWithValue("slovo", textBox4.Text);
                update_sotr.Parameters.AddWithValue("snils", maskedTextBox3.Text);
                update_sotr.Parameters.AddWithValue("dolzh_id", comboBox1.SelectedValue);
                update_sotr.ExecuteNonQuery();
                _CC.conection.Close();
                load_sotr();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source=DESKTOP-AG8AKLU;initial catalog=Zara_base;Persist Security info = True; User ID = SA; Password = qweqweqwe123");
            con.Open();
            SqlCommand delete_sotr = new SqlCommand("[dbo].delete_sotrud", con);
            delete_sotr.CommandType = CommandType.StoredProcedure;
            delete_sotr.Parameters.AddWithValue("@id_sotr", Convert.ToInt32(id_sotr));
            delete_sotr.ExecuteNonQuery();
            MessageBox.Show("Сотрудник удален");
            maskedTextBox9.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            maskedTextBox1.Text = "";
            maskedTextBox7.Text = "";
            maskedTextBox8.Text = "";
            maskedTextBox5.Text = "";
            maskedTextBox2.Text = "";
            maskedTextBox4.Text = "";
            maskedTextBox6.Text = "";
            maskedTextBox3.Text = "";
            comboBox1.Text = "";
            load_sotr();
            get_tabel();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++) //row-строка
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox5.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void главноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }
        private void ReplaceWordStub(string stud, string text, Word.Document wordDocument)
        {
            //string data = Convert.ToString(DateTime.Today);
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stud, ReplaceWith: text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
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
                    ReplaceWordStub("@fio", textBox1.Text + ' ' + textBox2.Text + ' ' + textBox3.Text, wordDociment);
                    ReplaceWordStub("@data", data, wordDociment);
                    ReplaceWordStub("@fio", textBox1.Text + ' ' + textBox2.Text + ' ' + textBox3.Text, wordDociment);
                    ReplaceWordStub("@seria", maskedTextBox7.Text, wordDociment);
                    ReplaceWordStub("@nomer", maskedTextBox8.Text, wordDociment);

                    //wordDociment.SaveAs(@"D:\docZARA\Трудовой");
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите сотрудника");
            }
        }


        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
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
                    ReplaceWordStub("@fio", textBox1.Text + ' ' + textBox2.Text + ' ' + textBox3.Text, wordDociment);
                    ReplaceWordStub("@tab", maskedTextBox9.Text, wordDociment);
                    ReplaceWordStub("@dol", comboBox1.Text, wordDociment);

                  // wordDociment.SaveAs(@"D:\docZARA\Трудовой");
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите сотрудника");
            }
        }

        private void kadr_sotr_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
