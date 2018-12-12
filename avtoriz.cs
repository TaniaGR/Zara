using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Data.SqlClient;

namespace Zara
{
    public partial class avtoriz : Form
    {
        podklclss _CC;
        string id_sotr;
        public avtoriz()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            glav_menu examp = new glav_menu();
            examp.Show();
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text == " -" || maskedTextBox2.Text == "")
            {
                MessageBox.Show("Не все поля заполнены!");
            }
            else
            {
                 _CC = new podklclss();
                 _CC.Set_Connection();
                 _CC.conection.Open();
                 SqlCommand newCMD = new SqlCommand("Select * from [dbo].[sotrud] where [login] ='" + maskedTextBox1.Text + "' and [parol] = '" + maskedTextBox2.Text + "'", _CC.conection);
                 SqlDataAdapter sda = new SqlDataAdapter(newCMD);
                 DataTable dt = new DataTable();
                 sda.Fill(dt);

                 if (dt.Rows.Count == 1)
                 {
                     Program.user_login = maskedTextBox1.Text;
                     MessageBox.Show("Вы успешно вошли в систему."  , "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     glav_menu glav_menu = new glav_menu();
                     this.Hide();
                     glav_menu.Show();

                 }
                 else
                 {
                     MessageBox.Show("Неправильный логин или пароль." + "\n" + "Повторите попытку входа!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 }
            }
        }

        private void avtoriz_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel2.Visible = true;
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (maskedTextBox5.Text == " -" || maskedTextBox4.Text == "")
            {
                MessageBox.Show("Не все поля заполнены!");
            }
            else
            {
                _CC = new podklclss();
                _CC.Set_Connection();
                _CC.conection.Open();
                SqlCommand newCMD = new SqlCommand("Select * from [dbo].[sotrud] where [login] ='" + maskedTextBox5.Text + "' and [slovo] = '" + maskedTextBox4.Text + "'", _CC.conection);
                SqlDataAdapter sda = new SqlDataAdapter(newCMD);
                DataTable dt = new DataTable();
                sda.Fill(dt);

                if (dt.Rows.Count == 1)
                {
                    
                    panel1.Visible = true;

                    _CC = new podklclss();
                    _CC.Set_Connection();
                    _CC.conection.Open();
                    SqlCommand cmd = new SqlCommand("Select * from [dbo].[sotrud](parol) where [login] ='" + maskedTextBox5.Text + _CC.conection);
                    
                    int result = ((int)cmd.ExecuteScalar());
                    _CC.conection.Close();
                    maskedTextBox3.Text = Convert.ToString(result);
                 
                                       
                }
                else
                {
                    MessageBox.Show("Неверные данные");
                }
            }
        }
    }
}
