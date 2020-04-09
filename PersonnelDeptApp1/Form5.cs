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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            //Вот тут необходимо после объединения модулей исправить подключение к БД
            String connectionString = "Server=hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com;User Id=postgres;Password=Ntcnbhjdfybt_01;Database=HRD;";
            NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            try
            {
                //Вот тут я запросом считываю из базы языки в comboBox
                string sqlExpression = "SELECT * FROM public.\"Language\"";
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
                //Вот тут я запросом считываю из базы степень владения языком в comboBox
                string sqlExpression1 = "SELECT * FROM public.\"DegreeLanguage\"";
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
        }

        //Обработка нажатия выход
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Обработка нажатия добавления языка
        private void button3_Click(object sender, EventArgs e)
        {
            //Добавляем язык и степень владения в datagridview на форме 1
            Program.f1.dataGridView3.Rows.Add();
            int kol = Program.f1.dataGridView3.Rows.Count - 1;
            Program.f1.dataGridView3.Rows[kol].Cells[0].Value = comboBox1.Text;
            Program.f1.dataGridView3.Rows[kol].Cells[1].Value = comboBox2.Text;
            MessageBox.Show("Язык добавлен в список!");
        }

        private void Form5_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Owner.Show();
        }
    }
}
