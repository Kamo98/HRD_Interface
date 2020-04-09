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
using Dadata;
using Dadata.Model;


namespace PersonnelDeptApp1
{
    public partial class Form1 : Form
    {
        SuggestClient api;
        
        string token = "a1bd42a8be5934f72b0a5802e26c61cd7458ac51";
        public int flag = 0; //Флаг добавления=0/редактирования=1
        public Form1()
        {
            InitializeComponent();
            init_dataGridExpFill();
        }

        private void init_dataGridExpFill() 
        {
            this.dataGridViewExpirience.Rows.Add(new object[] { "Дней"});
            this.dataGridViewExpirience.Rows.Add(new object[] { "Месяцев" });
            this.dataGridViewExpirience.Rows.Add(new object[] { "Лет" });
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        //Обработка нажатия кнопки применить
        private void button6_Click(object sender, EventArgs e)
        {
            //Вот тут необходимо после объединения модулей исправить подключение к БД
            String connectionString = "Server=hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com;User Id=postgres;Password=Ntcnbhjdfybt_01;Database=HRD;";
            NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            try
            {
                npgSqlConnection.Open();
               // MessageBox.Show("Подключение открыто!!");
               //Если форма открыта для добавления, делаем добавление
               if(flag==0)
                {
                    string surname = richTextBox14.Text;
                    //Если фамилия не заполнена
                    if (surname == "")
                    {
                        MessageBox.Show("Введите пожалуйста фамилию сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string name = richTextBox15.Text;
                    //Если имя не заполнено
                    if (name == "")
                    {
                        MessageBox.Show("Введите пожалуйста имя сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string otchestvo = richTextBox13.Text;
                    //Если отчество не заполнено
                    if (otchestvo == "")
                    {
                        MessageBox.Show("Введите пожалуйста отчество сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    DateTime Date_birth = dateTimePicker6.Value;

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

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel18_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel17_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel15_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label71_Click(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        //Загрузка формы
        private void Form1_Load(object sender, EventArgs e)
        {
            //Вот тут необходимо после объединения модулей исправить подключение к БД
            String connectionString = "Server=hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com;User Id=postgres;Password=Ntcnbhjdfybt_01;Database=HRD;";
            NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            try
            {
                //Вот тут я запросом считываю из базы Гражданство в comboBox
                string sqlExpression = "SELECT * FROM public.\"Citizenship\"";
                npgSqlConnection.Open();
                // MessageBox.Show("Подключение открыто!!");
                NpgsqlCommand command = new NpgsqlCommand(sqlExpression, npgSqlConnection);
                NpgsqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows) // если есть данные
                {
                    while (reader.Read()) // построчно считываем данные
                    {
                        object Name = reader.GetValue(1);
                        comboBox9.Items.Add(Name);
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
                //Вот тут я запросом считываю из базы состояние в браке в comboBox
                string sqlExpression1 = "SELECT * FROM public.\"MaritalStatus\"";
                npgSqlConnection.Open();
                NpgsqlCommand command1 = new NpgsqlCommand(sqlExpression1, npgSqlConnection);
                NpgsqlDataReader reader1 = command1.ExecuteReader();
                if (reader1.HasRows) // если есть данные
                {
                    while (reader1.Read()) // построчно считываем данные
                    {
                        object Name = reader1.GetValue(1);
                        comboBox7.Items.Add(Name);
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
                //Вот тут я запросом считываю из базы характер работы в comboBox
                string sqlExpression2 = "SELECT * FROM public.\"CharacterWork\"";
                npgSqlConnection.Open();
                NpgsqlCommand command2 = new NpgsqlCommand(sqlExpression2, npgSqlConnection);
                NpgsqlDataReader reader2 = command2.ExecuteReader();
                if (reader2.HasRows) // если есть данные
                {
                    while (reader2.Read()) // построчно считываем данные
                    {
                        object Name = reader2.GetValue(1);
                        comboBox4.Items.Add(Name);
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

        private void tableLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label41_Click(object sender, EventArgs e)
        {

        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel17_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel17_Paint_2(object sender, PaintEventArgs e)
        {

        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void richTextBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label42_Click(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void label44_Click(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click_1(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridViewExpirience_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form ifrm = Application.OpenForms[0];
            ifrm.Show();
        }

        //Ограничение на ввод серии и номера паспорта
        private void richTextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ограничение ввода можем ввести только цифры для серии и номера и пробел между ними
            if (e.KeyChar < '0' | e.KeyChar > '9' && e.KeyChar != (char)Keys.Back && e.KeyChar != ' ')
            {

                e.Handled = true;
            }
            //ограничение пробел можно ставить только между серией и номером паспорта
            if (richTextBox6.SelectionStart != 4 & e.KeyChar == ' ')
            {
                e.Handled = true;
            }
            //ограничение серия 4 цифры номер 6
            if (richTextBox6.Text.Length > 10)
            {
                e.Handled = true;
            }
        }

        //Ограничение на ввод индекса в адресе по паспорту
        private void richTextBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ограничение ввода можем ввести только цифры
            if (e.KeyChar < '0' | e.KeyChar > '9' && e.KeyChar != (char)Keys.Back)
            {

                e.Handled = true;
            }
            
            //ограничение на 6 знаков
            if (richTextBox8.Text.Length > 5)
            {
                e.Handled = true;
            }
        }

        private void richTextBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ограничение ввода можем ввести только цифры
            if (e.KeyChar < '0' | e.KeyChar > '9' && e.KeyChar != (char)Keys.Back)
            {

                e.Handled = true;
            }

            //ограничение на 6 знаков
            if (richTextBox10.Text.Length > 5)
            {
                e.Handled = true;
            }
        }

        private void tableLayoutPanel17_Paint_3(object sender, PaintEventArgs e)
        {

        }

        private void richTextBox7_TextChanged(object sender, EventArgs e)
        {
            var api = new SuggestClient(token);

            AutoCompleteStringCollection help = new AutoCompleteStringCollection();
            var response = api.SuggestAddress(richTextBox7.Text);

            listBox1.Items.Clear();
            for (int i = 0; i < response.suggestions.Count; i++)
            {
                listBox1.Items.Add(response.suggestions[i].value.ToString());
            }
        }

        //Ограничение ввода ИНН
        private void richTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ограничение ввода можем ввести только цифры
            if (e.KeyChar < '0' | e.KeyChar > '9' && e.KeyChar != (char)Keys.Back)
            {

                e.Handled = true;
            }

            //ограничение на 12 знаков
            if (richTextBox4.Text.Length > 11)
            {
                e.Handled = true;
            }
        }

        //Ограничение ввода страхового свидетельства
        private void richTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ограничение ввода можем ввести только цифры
            if (e.KeyChar < '0' | e.KeyChar > '9' && e.KeyChar != (char)Keys.Back && e.KeyChar!='-')
            {

                e.Handled = true;
            }

            //ограничение на 12 знаков
            if (richTextBox4.Text.Length > 13)
            {
                e.Handled = true;
            }
          
        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            richTextBox7.Text = listBox1.SelectedItem.ToString();
        }
    }
}
