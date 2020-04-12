using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;

namespace PersonnelDeptApp1
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }

        //Обработка нажатия выход
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form6_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Owner.Show();
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            //Вот тут необходимо после объединения модулей исправить подключение к БД
            String connectionString = "Server=hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com;User Id=postgres;Password=Ntcnbhjdfybt_01;Database=HRD;";
            NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            try
            {
                //Вот тут я запросом считываю из базы тип образования в comboBox
                string sqlExpression = "SELECT * FROM public.\"Education\"";
                npgSqlConnection.Open();
                // MessageBox.Show("Подключение открыто!!");
                NpgsqlCommand command = new NpgsqlCommand(sqlExpression, npgSqlConnection);
                NpgsqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows) // если есть данные
                {
                    while (reader.Read()) // построчно считываем данные
                    {
                        object Name = reader.GetValue(1);
                        comboBox1.Items.Add(Name);
                    }
                }

            }
            catch (NpgsqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                npgSqlConnection.Close();
                //  MessageBox.Show("Подключение закрыто!!");
            }

            try
            {
                //Вот тут я запросом считываю из базы образовательное учереждение в comboBox
                string sqlExpression1 = "SELECT * FROM public.\"Institution\"";
                npgSqlConnection.Open();
                // MessageBox.Show("Подключение открыто!!");
                NpgsqlCommand command1 = new NpgsqlCommand(sqlExpression1, npgSqlConnection);
                NpgsqlDataReader reader1 = command1.ExecuteReader();
                if (reader1.HasRows) // если есть данные
                {
                    while (reader1.Read()) // построчно считываем данные
                    {
                        object Name = reader1.GetValue(1);
                        comboBox2.Items.Add(Name);
                    }
                }

            }
            catch (NpgsqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                npgSqlConnection.Close();
                //  MessageBox.Show("Подключение закрыто!!");
            }

            try
            {
                //Вот тут я запросом считываю из базы профиль подготовки в comboBox
                string sqlExpression2 = "SELECT * FROM public.\"Specialty\"";
                npgSqlConnection.Open();
                // MessageBox.Show("Подключение открыто!!");
                NpgsqlCommand command2 = new NpgsqlCommand(sqlExpression2, npgSqlConnection);
                NpgsqlDataReader reader2 = command2.ExecuteReader();
                if (reader2.HasRows) // если есть данные
                {
                    while (reader2.Read()) // построчно считываем данные
                    {
                        object Name = reader2.GetValue(1);
                        comboBox3.Items.Add(Name);
                    }
                }

            }
            catch (NpgsqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                npgSqlConnection.Close();
                //  MessageBox.Show("Подключение закрыто!!");
            }
        }

        //Ограничение ввода года
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ограничение ввода можем ввести только цифры
            if (e.KeyChar < '0' | e.KeyChar > '9' && e.KeyChar != (char)Keys.Back)
            {

                e.Handled = true;
            }

            //ограничение на 12 знаков
            if (textBox3.Text.Length > 3 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        //Обработка нажатия добавить образование
        private void button11_Click(object sender, EventArgs e)
        {
            //Сначала проверим, что все поля заполнеы
            if (textBox1.Text == "")
            {
                MessageBox.Show("Введите пожалуйста наименование документа об образовании!");
                return;
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("Введите пожалуйста серийный номер!");
                return;
            }
            if (textBox3.Text == "")
            {
                MessageBox.Show("Введите пожалуйста год!");
                return;
            }
            if (string.IsNullOrEmpty(comboBox1.Text))
            {
                MessageBox.Show("Не выбран тип образования!");
                return;
            }
            if (string.IsNullOrEmpty(comboBox2.Text))
            {
                MessageBox.Show("Не выбрано образовательное учереждение!");
                return;
            }
            if (string.IsNullOrEmpty(comboBox3.Text))
            {
                MessageBox.Show("Не выбран профиль подготовки!");
                return;
            }
            //Добавляем все пункты образования в datagridview на форме 1
            Program.f1.dataGridView4.Rows.Add();
            int kol = Program.f1.dataGridView4.Rows.Count - 1;
            Program.f1.dataGridView4.Rows[kol].Cells[0].Value = textBox1.Text;
            Program.f1.dataGridView4.Rows[kol].Cells[1].Value = comboBox1.Text;
            Program.f1.dataGridView4.Rows[kol].Cells[2].Value = comboBox2.Text;
            Program.f1.dataGridView4.Rows[kol].Cells[3].Value = comboBox2.Text;
            Program.f1.dataGridView4.Rows[kol].Cells[4].Value = textBox2.Text;
            Program.f1.dataGridView4.Rows[kol].Cells[5].Value = textBox3.Text;
            MessageBox.Show("Образование добавлено в список!");
        }
    }
}
