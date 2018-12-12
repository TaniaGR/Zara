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
    public partial class prodant : Form
    {
        podklclss _CC;
        public prodant()
        {
            InitializeComponent();
        }

        private void prod_tov()
        {
            _CC = new podklclss();
            _CC.Set_Connection();
            SqlCommand tov_prod = new SqlCommand("select * from prod_tov", _CC.conection);
            _CC.conection.Open();
            SqlDataReader tov = tov_prod.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(tov);
            dataGridView1.DataSource = dt;            
            dataGridView1.Columns[0].HeaderText = "Артикул";
            dataGridView1.Columns[1].HeaderText = "Наименование товара";
            dataGridView1.Columns[2].HeaderText = "Количество";
            dataGridView1.Columns[3].HeaderText = "Цена за шт";
            dataGridView1.Columns[4].HeaderText = "Номер чека";

            _CC.conection.Close();
        }
        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            prodazhi examp = new prodazhi();
            examp.Show();
            this.Close();
        }

        private void вернутьсяВГлавноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }

        private void prodant_Load(object sender, EventArgs e)
        {
            prod_tov();
        }

        private void главноеМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++) //row-строка
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void prodant_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
