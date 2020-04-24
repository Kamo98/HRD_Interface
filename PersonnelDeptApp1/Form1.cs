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
using System.IO;


namespace PersonnelDeptApp1
{
    public partial class Form1 : Form
    {
        SuggestClient api;
        string token = "a1bd42a8be5934f72b0a5802e26c61cd7458ac51";
        public int pk_personal_card;
        public int flag = 0; //Флаг добавления=0/редактирования=1
        public int flagnewfile = -1; //Флаг для файлов характеристики
        public Form1()
        {
            Program.f1 = this;
            InitializeComponent();
            init_dataGridExpFill();
        }

		public Form1(int id_personal_card)
		{
			Program.f1 = this;
			InitializeComponent();
			init_dataGridExpFill();

			pk_personal_card = id_personal_card;
			flag = 1;
		}


		private void init_dataGridExpFill() 
        {
            //this.dataGridViewExpirience.Rows.Add(new object[] { "Дней"});
            //this.dataGridViewExpirience.Rows.Add(new object[] { "Месяцев" });
            //this.dataGridViewExpirience.Rows.Add(new object[] { "Лет" });
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        //Обработка нажатия кнопки применить
        private void button6_Click(object sender, EventArgs e)
        {
            //Вот тут необходимо после объединения модулей исправить подключение к БД
            String connectionString = "Server=hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com;User Id=" + Connection.login + ";Password=" + Connection.password + ";Database=HRD;";
            NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            try
            {
                npgSqlConnection.Open();
               // MessageBox.Show("Подключение открыто!!");
               
                    //Проверяем, что все необходимые поля заполнены
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
                    string Character_work = comboBox4.Text;
                    int pk_character_work=-1;
                    //Если характер работы не заполнен
                    if (string.IsNullOrEmpty(comboBox4.Text))
                    {
                        MessageBox.Show("Не выбран характер работы!");
                        return;
                    }
                    else
                    {
                        string SqlExpression = "SELECT * FROM public.\"CharacterWork\" WHERE \"Name\"=@Character_work";
                        NpgsqlConnection npgSqlConnection1 = new NpgsqlConnection(connectionString);
                        npgSqlConnection1.Open();
                        using (npgSqlConnection1)
                        {
                            
                            NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection1);
                            // создаем параметр для имени
                            NpgsqlParameter CWParam = new NpgsqlParameter("@Character_work", Character_work);
                            command.Parameters.Add(CWParam);
                            NpgsqlDataReader reader = command.ExecuteReader();
                            if (reader.HasRows) // если есть данные
                            {
                                while (reader.Read()) // построчно считываем данные
                                {
                                    object character_work = reader.GetValue(0);
                                    pk_character_work = Convert.ToInt32(character_work);
                                    
                                }
                            }
                            

                        }
                        npgSqlConnection1.Close();
                        // MessageBox.Show(pk_character_work.ToString());
                    }
                    DateTime Date_birth = dateTimePicker6.Value;
                    string Marital_status= comboBox7.Text;
                    int pk_marital_status = -1;
                    if (string.IsNullOrEmpty(comboBox7.Text))
                    {
                        MessageBox.Show("Не выбрано состояние в браке!");
                        return;
                    }
                    else
                    {
                        string SqlExpression = "SELECT * FROM public.\"MaritalStatus\" WHERE \"Name\"=@Marital_status";
                        NpgsqlConnection npgSqlConnection2 = new NpgsqlConnection(connectionString);
                        npgSqlConnection2.Open();
                        using (npgSqlConnection2)
                        {
                            
                            NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection2);
                            // создаем параметр для имени
                            NpgsqlParameter msParam = new NpgsqlParameter("@Marital_status", Marital_status);
                            command.Parameters.Add(msParam);
                            NpgsqlDataReader reader1 = command.ExecuteReader();
                            if (reader1.HasRows) // если есть данные
                            {
                                while (reader1.Read()) // построчно считываем данные
                                {
                                    object martial_status = reader1.GetValue(0);
                                    pk_marital_status = Convert.ToInt32(martial_status);

                                }
                            }


                        }
                        npgSqlConnection2.Close();
                        //MessageBox.Show(pk_marital_status.ToString());
                    }
                    string Citizenship= comboBox9.Text;
                    int pk_citizenship = -1;
                    if (string.IsNullOrEmpty(comboBox9.Text))
                    {
                        MessageBox.Show("Не выбрано гражданство!");
                        return;
                    }
                    else
                    {
                        string SqlExpression = "SELECT * FROM public.\"Citizenship\" WHERE \"Name\"=@Citizenship";
                        NpgsqlConnection npgSqlConnection3 = new NpgsqlConnection(connectionString);
                        npgSqlConnection3.Open();
                        using (npgSqlConnection3)
                        {

                            NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection3);
                            // создаем параметр для имени
                            NpgsqlParameter CWParam = new NpgsqlParameter("@Citizenship", Citizenship);
                            command.Parameters.Add(CWParam);
                            NpgsqlDataReader reader = command.ExecuteReader();
                            if (reader.HasRows) // если есть данные
                            {
                                while (reader.Read()) // построчно считываем данные
                                {
                                    object citizenship = reader.GetValue(0);
                                    pk_citizenship = Convert.ToInt32(citizenship);

                                }
                            }


                        }
                        npgSqlConnection3.Close();
                        //MessageBox.Show(pk_citizenship.ToString());
                    }
                    string INN = richTextBox4.Text;
                    //Если ИНН не заполнено
                    if (INN == "")
                    {
                        MessageBox.Show("Введите пожалуйста ИНН сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    else if (INN.Length != 12)
                    {
                        MessageBox.Show("ИНН должен состоять из 12 цифр!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string SNN = richTextBox3.Text;
                    //Если номер страхового свидетельства не заполнен
                    if (SNN == "")
                    {
                        MessageBox.Show("Введите пожалуйста номер страхового свидетельства сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    else if (SNN.Length != 14)
                    {
                        MessageBox.Show("Номер страхового свидетельства должен состоять из 14 символов!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string Serial_number = richTextBox6.Text;
                    //Если серия и номер паспорта не заполнены
                    if (Serial_number == "")
                    {
                        MessageBox.Show("Введите пожалуйста серию и номер паспорта сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    //Тут такая проверка, потому что между серией и номером можно поставить пробел, а можно не ставить
                    else if (Serial_number.Length != 11 && Serial_number.Length != 10)
                    {
                        MessageBox.Show("Серия 4 цифры и номер паспорта 6!");
                        npgSqlConnection.Close();
                        return;
                    }
                    DateTime Passport_date = dateTimePicker4.Value;
                    string Vidan = richTextBox5.Text;
                    //Если кем выдан паспорт не заполнено
                    if (Vidan == "")
                    {
                        MessageBox.Show("Введите пожалуйста кем выдан паспорт сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string Index_real = richTextBox8.Text;
                    //Если индекс прописки не заполнен
                    if (Index_real == "")
                    {
                        MessageBox.Show("Введите пожалуйста индекс прописки сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    else if (Index_real.Length!=6)
                    {
                        MessageBox.Show("Индекс прописки должен состоять из 6 цифр!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string Propiska = richTextBox9.Text;
                    //Если адрес прописки не заполнен
                    if (Propiska == "")
                    {
                        MessageBox.Show("Введите пожалуйста адрес прописки сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    DateTime Home_date = dateTimePicker1.Value;
                    //Фактический адрес считываю без проверки, он может быть не заполнен
                    string Index_fact = richTextBox10.Text;
                    if (Index_fact.Length != 6 && Index_fact!="")
                    {
                        MessageBox.Show("Индекс фактического адреса должен состоять из 6 цифр!");
                        npgSqlConnection.Close();
                        return;
                    }
                    string Fact_address = richTextBox11.Text;
                    //Телефон без проверки, так же может быть не заполнен
                    string Phone= richTextBox2.Text;
                    string Birth_place = richTextBox7.Text;
                    //Если место рождения не заполнено
                    if (Birth_place == "")
                    {
                        MessageBox.Show("Введите пожалуйста место рождения сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }
                    DateTime Creation_date = dateTimePicker2.Value;
                    String Gender;
                    if (radioButton3.Checked==true)
                    {
                        Gender = "М";
                    }
                    else
                    {
                        Gender = "Ж";
                    }
                    string Work_kind = richTextBox12.Text;
                    //Если вид работы не заполнен
                    if (Work_kind == "")
                    {
                        MessageBox.Show("Введите пожалуйста вид работы сотрудника!");
                        npgSqlConnection.Close();
                        return;
                    }

                    string Military_profile = richTextBox16.Text;
                    string Military_code = richTextBox17.Text;
                    string Military_name= richTextBox18.Text;
                    string Military_status = comboBox11.Text;
                    string Military_cancel= richTextBox19.Text;

                    string military_rank = comboBox6.Text;
                    int pk_military_rank = 1;
                    if (string.IsNullOrEmpty(comboBox6.Text))
                    {
                        MessageBox.Show("Не выбрано воинское звание");
                        return;
                    }
                    else
                    {
                        string SqlExpression = "SELECT * FROM public.\"MilitaryRank\" WHERE \"Name\"=@military_rank";
                        NpgsqlConnection npgSqlConnection4 = new NpgsqlConnection(connectionString);
                        npgSqlConnection4.Open();
                        using (npgSqlConnection4)
                        {

                            NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection4);
                            // создаем параметр для имени
                            NpgsqlParameter CWParam = new NpgsqlParameter("@military_rank", military_rank);
                            command.Parameters.Add(CWParam);
                            NpgsqlDataReader reader = command.ExecuteReader();
                            if (reader.HasRows) // если есть данные
                            {
                                while (reader.Read()) // построчно считываем данные
                                {
                                    object military_rank1 = reader.GetValue(0);
                                    pk_military_rank = Convert.ToInt32(military_rank1);

                                }
                            }


                        }
                        npgSqlConnection4.Close();
                        //MessageBox.Show(pk_military_rank.ToString());
                    }
                    string stock_category = comboBox8.Text;
                    int pk_stock_category = -1;
                    if (string.IsNullOrEmpty(comboBox8.Text))
                    {
                        MessageBox.Show("Не выбрана категория запаса");
                        return;
                    }
                    else
                    {
                        string SqlExpression = "SELECT * FROM public.\"StockCategory\" WHERE \"Name\"=@stock_category";
                        NpgsqlConnection npgSqlConnection5 = new NpgsqlConnection(connectionString);
                        npgSqlConnection5.Open();
                        using (npgSqlConnection5)
                        {

                            NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection5);
                            // создаем параметр для имени
                            NpgsqlParameter CWParam = new NpgsqlParameter("@stock_category", stock_category);
                            command.Parameters.Add(CWParam);
                            NpgsqlDataReader reader = command.ExecuteReader();
                            if (reader.HasRows) // если есть данные
                            {
                                while (reader.Read()) // построчно считываем данные
                                {
                                    object stock_category1 = reader.GetValue(0);
                                    pk_stock_category = Convert.ToInt32(stock_category1);

                                }
                            }


                        }
                        npgSqlConnection5.Close();
                        //MessageBox.Show(pk_stock_category.ToString());
                    }
                    string category_military = comboBox5.Text;
                    int pk_category_military = -1;
                    if (string.IsNullOrEmpty(comboBox5.Text))
                    {
                        MessageBox.Show("Не выбрана категория годности");
                        return;
                    }
                    else
                    {
                        string SqlExpression = "SELECT * FROM public.\"CategoryMilitary\" WHERE \"Name\"=@category_military";
                        NpgsqlConnection npgSqlConnection6 = new NpgsqlConnection(connectionString);
                        npgSqlConnection6.Open();
                        using (npgSqlConnection6)
                        {

                            NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection6);
                            // создаем параметр для имени
                            NpgsqlParameter CWParam = new NpgsqlParameter("@category_military", category_military);
                            command.Parameters.Add(CWParam);
                            NpgsqlDataReader reader = command.ExecuteReader();
                            if (reader.HasRows) // если есть данные
                            {
                                while (reader.Read()) // построчно считываем данные
                                {
                                    object category_military1 = reader.GetValue(0);
                                    pk_category_military = Convert.ToInt32(category_military1);

                                }
                            }


                        }
                        npgSqlConnection6.Close();
                       // MessageBox.Show(pk_category_military.ToString());
                    }

                    
                    string Characteristic = richTextBox1.Text;

                //Если форма открыта для добавления добавляем нового сотрудника
                if (flag == 0)
                {
                    string sqlExpression = "INSERT INTO \"PersonalCard\" (\"pk_marital_status\",\"pk_character_work\",\"surname\"," +
                        "\"name\",\"otchestvo\",\"birthday\",\"Characteristic\",\"INN\",\"SSN\",\"Serial_number\"" +
                        ",\"Passport_date\",\"Vidan\",\"Home_date\",\"Propiska\",\"Fact_address\",\"Phone\",\"pk_military_rank\",\"pk_category_military\"," +
                        "\"pk_stock_category\",\"Birth_place\",\"Creation_date\",\"Gender\",\"Military_profile\",\"Military_code\",\"Military_name\"" +
                        ",\"Military_status\",\"Military_cancel\",\"Work_kind\",\"Index_fact\",\"Index_real\") " +
                        "VALUES (@pk_marital_status,@pk_character_work,@surname,@name,@otchestvo,@birthday," +
                        "@Characteristic,@INN,@SSN,@Serial_number,@Passport_date,@Vidan,@Home_date,@Propiska,@Fact_address," +
                        "@Phone,@pk_military_rank,@pk_category_military,@pk_stock_category,@Birth_place,@Creation_date,@Gender,@Military_profile," +
                        "@Military_code,@Military_name,@Military_status,@Military_cancel,@Work_kind,@Index_fact,@Index_real) RETURNING \"pk_personal_card\"";
                    using (npgSqlConnection)
                    { 
                        NpgsqlCommand command = new NpgsqlCommand(sqlExpression, npgSqlConnection);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_marital_status", pk_marital_status);
                        command.Parameters.Add(Param1);
                        NpgsqlParameter Param2 = new NpgsqlParameter("@pk_character_work", pk_character_work);
                        command.Parameters.Add(Param2);
                        NpgsqlParameter Param3 = new NpgsqlParameter("@surname", surname);
                        command.Parameters.Add(Param3);
                        NpgsqlParameter Param4 = new NpgsqlParameter("@name", name);
                        command.Parameters.Add(Param4);
                        NpgsqlParameter Param5 = new NpgsqlParameter("@otchestvo", otchestvo);
                        command.Parameters.Add(Param5);
                        NpgsqlParameter Param6 = new NpgsqlParameter("@birthday", Date_birth);
                        command.Parameters.Add(Param6);
                        NpgsqlParameter Param7 = new NpgsqlParameter("@Characteristic", Characteristic);
                        command.Parameters.Add(Param7);
                        NpgsqlParameter Param8 = new NpgsqlParameter("@INN", INN);
                        command.Parameters.Add(Param8);
                        NpgsqlParameter Param9 = new NpgsqlParameter("@SSN", SNN);
                        command.Parameters.Add(Param9);
                        NpgsqlParameter Param10 = new NpgsqlParameter("@Serial_number", Serial_number);
                        command.Parameters.Add(Param10);
                        NpgsqlParameter Param11 = new NpgsqlParameter("@Passport_date", Passport_date);
                        command.Parameters.Add(Param11);
                        NpgsqlParameter Param12 = new NpgsqlParameter("@Vidan", Vidan);
                        command.Parameters.Add(Param12);
                        NpgsqlParameter Param13 = new NpgsqlParameter("@Home_date", Home_date);
                        command.Parameters.Add(Param13);
                        NpgsqlParameter Param14 = new NpgsqlParameter("@Propiska", Propiska);
                        command.Parameters.Add(Param14);
                        NpgsqlParameter Param15 = new NpgsqlParameter("@Fact_address", Fact_address);
                        command.Parameters.Add(Param15);
                        NpgsqlParameter Param16 = new NpgsqlParameter("@Phone", Phone);
                        command.Parameters.Add(Param16);
                        NpgsqlParameter Param17 = new NpgsqlParameter("@pk_military_rank", pk_military_rank);
                        command.Parameters.Add(Param17);
                        NpgsqlParameter Param18 = new NpgsqlParameter("@pk_category_military", pk_category_military);
                        command.Parameters.Add(Param18);
                        NpgsqlParameter Param19 = new NpgsqlParameter("@pk_stock_category", pk_stock_category);
                        command.Parameters.Add(Param19);
                        NpgsqlParameter Param20 = new NpgsqlParameter("@Birth_place", Birth_place);
                        command.Parameters.Add(Param20);
                        NpgsqlParameter Param21 = new NpgsqlParameter("@Creation_date", Creation_date);
                        command.Parameters.Add(Param21);
                        NpgsqlParameter Param22 = new NpgsqlParameter("@Gender", Gender);
                        command.Parameters.Add(Param22);
                        NpgsqlParameter Param23 = new NpgsqlParameter("@Military_profile", Military_profile);
                        command.Parameters.Add(Param23);
                        NpgsqlParameter Param24 = new NpgsqlParameter("@Military_code", Military_code);
                        command.Parameters.Add(Param24);
                        NpgsqlParameter Param25 = new NpgsqlParameter("@Military_name", Military_name);
                        command.Parameters.Add(Param25);
                        NpgsqlParameter Param26 = new NpgsqlParameter("@Military_status", Military_status);
                        command.Parameters.Add(Param26);
                        NpgsqlParameter Param27 = new NpgsqlParameter("@Military_cancel", Military_cancel);
                        command.Parameters.Add(Param27);
                        NpgsqlParameter Param28 = new NpgsqlParameter("@Work_kind", Work_kind);
                        command.Parameters.Add(Param28);
                        NpgsqlParameter Param29 = new NpgsqlParameter("@Index_fact", Index_fact);
                        command.Parameters.Add(Param29);
                        NpgsqlParameter Param30 = new NpgsqlParameter("@Index_real", Index_real);
                        command.Parameters.Add(Param30);
                        NpgsqlDataReader reader2 = command.ExecuteReader();
                        if (reader2.HasRows) // если есть данные
                        {
                            while (reader2.Read()) // построчно считываем данные
                            {
                                object pk = reader2.GetValue(0);
                                pk_personal_card = Convert.ToInt32(pk);

                            }
                        }

                        //int number = command.ExecuteNonQuery();

                        //MessageBox.Show("Сотрудник добавлен успешно!");
                        
                        
                    }
                    //Проверка, были ли добавлены языки. Если да, то добавляем информацию в базу
                    if(dataGridView3.Rows.Count != 0) 
                    {
                        foreach (DataGridViewRow Row in dataGridView3.Rows)
                        {
                            string language = Row.Cells[0].Value.ToString();
                            string language_degree = Row.Cells[1].Value.ToString();
                            int pk_language = -1;
                            int pk_degree_language = -1;
                            string SqlExpression = "SELECT * FROM public.\"Language\" WHERE \"Name\"=@language";
                            NpgsqlConnection npgSqlConnection7 = new NpgsqlConnection(connectionString);
                            npgSqlConnection7.Open();
                            using (npgSqlConnection7)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection7);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@language", language);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader7 = command.ExecuteReader();
                                if (reader7.HasRows) // если есть данные
                                {
                                    while (reader7.Read()) // построчно считываем данные
                                    {
                                        object language1 = reader7.GetValue(0);
                                        pk_language = Convert.ToInt32(language1);

                                    }
                                }


                            }
                            npgSqlConnection7.Close();

                            string SqlExpression1 = "SELECT * FROM public.\"DegreeLanguage\" WHERE \"Name\"=@language_degree";
                            NpgsqlConnection npgSqlConnection8 = new NpgsqlConnection(connectionString);
                            npgSqlConnection8.Open();
                            using (npgSqlConnection8)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression1, npgSqlConnection8);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@language_degree", language_degree);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader8 = command.ExecuteReader();
                                if (reader8.HasRows) // если есть данные
                                {
                                    while (reader8.Read()) // построчно считываем данные
                                    {
                                        object degree_language1 = reader8.GetValue(0);
                                        pk_degree_language = Convert.ToInt32(degree_language1);

                                    }
                                }


                            }
                            npgSqlConnection8.Close();

                            string SqlExpression2 = "INSERT INTO \"lang-card\" (\"pk_language\",\"pk_personal_card\",\"pk_degree_language\") " +
                                "VALUES (@pk_language,@pk_personal_card,@pk_degree_language)";
                            NpgsqlConnection npgSqlConnection9 = new NpgsqlConnection(connectionString);
                            npgSqlConnection9.Open();
                            using (npgSqlConnection9)
                            {
                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression2, npgSqlConnection9);
                                // создаем параметры и добавляем их к команде
                                NpgsqlParameter Param1 = new NpgsqlParameter("@pk_language", pk_language);
                                command.Parameters.Add(Param1);
                                NpgsqlParameter Param2 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                                command.Parameters.Add(Param2);
                                NpgsqlParameter Param3 = new NpgsqlParameter("@pk_degree_language", pk_degree_language);
                                command.Parameters.Add(Param3);
                                int number = command.ExecuteNonQuery();
                                //MessageBox.Show("Знание языков добавлено успешно!");

                            }
                            npgSqlConnection9.Close();
                        }
                    }
                    //Проверка, были ли добавлены образования. Если да, то добавляем информацию в базу
                    if (dataGridView4.Rows.Count != 0)
                    {
                        foreach (DataGridViewRow Row in dataGridView4.Rows)
                        {
                            string document_name = Row.Cells[0].Value.ToString();
                            string education = Row.Cells[1].Value.ToString();
                            string institution = Row.Cells[2].Value.ToString();
                            string specialty = Row.Cells[3].Value.ToString();
                            string serial_number = Row.Cells[4].Value.ToString();
                            int Year = Convert.ToInt32(Row.Cells[5].Value);
                            int pk_education = -1;
                            int pk_specialty = -1;
                            int pk_nstitution = -1;
                            string SqlExpression = "SELECT * FROM public.\"Education\" WHERE \"Name\"=@education";
                            NpgsqlConnection npgSqlConnection10 = new NpgsqlConnection(connectionString);
                            npgSqlConnection10.Open();
                            using (npgSqlConnection10)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection10);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@education", education);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader10 = command.ExecuteReader();
                                if (reader10.HasRows) // если есть данные
                                {
                                    while (reader10.Read()) // построчно считываем данные
                                    {
                                        object education1 = reader10.GetValue(0);
                                        pk_education = Convert.ToInt32(education1);

                                    }
                                }


                            }
                            npgSqlConnection10.Close();
                            string SqlExpression1 = "SELECT * FROM public.\"Specialty\" WHERE \"Name\"=@specialty";
                            NpgsqlConnection npgSqlConnection11 = new NpgsqlConnection(connectionString);
                            npgSqlConnection11.Open();
                            using (npgSqlConnection11)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression1, npgSqlConnection11);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@specialty", specialty);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader11 = command.ExecuteReader();
                                if (reader11.HasRows) // если есть данные
                                {
                                    while (reader11.Read()) // построчно считываем данные
                                    {
                                        object speciality1 = reader11.GetValue(0);
                                        pk_specialty = Convert.ToInt32(speciality1);

                                    }
                                }


                            }
                            npgSqlConnection11.Close();
                            string SqlExpression2 = "SELECT * FROM public.\"Institution\" WHERE \"Name\"=@institution";
                            NpgsqlConnection npgSqlConnection12 = new NpgsqlConnection(connectionString);
                            npgSqlConnection12.Open();
                            using (npgSqlConnection12)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression2, npgSqlConnection12);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@institution", institution);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader12 = command.ExecuteReader();
                                if (reader12.HasRows) // если есть данные
                                {
                                    while (reader12.Read()) // построчно считываем данные
                                    {
                                        object institution1 = reader12.GetValue(0);
                                        pk_nstitution = Convert.ToInt32(institution1);

                                    }
                                }


                            }
                            npgSqlConnection12.Close();
                            string SqlExpression3 = "INSERT INTO \"card-education\" (\"pk_education\",\"pk_personal_card\",\"pk_specialty\"," +
                                "\"pk_nstitution\",\"document_name\",\"serial_number\",\"Year\") " +
                                "VALUES (@pk_education,@pk_personal_card,@pk_specialty,@pk_nstitution,@document_name,@serial_number,@Year)";
                            NpgsqlConnection npgSqlConnection13 = new NpgsqlConnection(connectionString);
                            npgSqlConnection13.Open();
                            using (npgSqlConnection13)
                            {
                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression3, npgSqlConnection13);
                                // создаем параметры и добавляем их к команде
                                NpgsqlParameter Param1 = new NpgsqlParameter("@pk_education", pk_education);
                                command.Parameters.Add(Param1);
                                NpgsqlParameter Param2 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                                command.Parameters.Add(Param2);
                                NpgsqlParameter Param3 = new NpgsqlParameter("@pk_specialty", pk_specialty);
                                command.Parameters.Add(Param3);
                                NpgsqlParameter Param4 = new NpgsqlParameter("@pk_nstitution", pk_nstitution);
                                command.Parameters.Add(Param4);
                                NpgsqlParameter Param5 = new NpgsqlParameter("@document_name", document_name);
                                command.Parameters.Add(Param5);
                                NpgsqlParameter Param6 = new NpgsqlParameter("@serial_number", serial_number);
                                command.Parameters.Add(Param6);
                                NpgsqlParameter Param7 = new NpgsqlParameter("@Year", Year);
                                command.Parameters.Add(Param7);
                                int number = command.ExecuteNonQuery();
                                //MessageBox.Show("Образование добавлено успешно!");

                            }
                            npgSqlConnection13.Close();
                        }
                    }

                    //Проверка, были ли добавлены характеристики. Если да, то добавляем информацию в базу
                    if (dataGridView5.Rows.Count != 0)
                    {
                        foreach (DataGridViewRow Row in dataGridView5.Rows)
                        {
                            string date = Row.Cells[0].Value.ToString();
                            DateTime date_formate = Convert.ToDateTime(date);
                            string fileReference = Row.Cells[1].Value.ToString();
                            string characteristic_date;
                            FileStream stream = new FileStream(Row.Cells[1].Value.ToString(), FileMode.Open);
                            StreamReader reader = new StreamReader(stream);
                            characteristic_date = (reader.ReadToEnd());
                            richTextBox1.Text = (reader.ReadToEnd());
                            reader.Close();
                            string SqlExpression4 = "INSERT INTO \"Characteristic\" (\"date\",\"characteristic\",\"fileReference\"," +
                                "\"pk_personal_card\") " +
                                "VALUES (@date,@characteristic,@fileReference,@pk_personal_card)";
                            NpgsqlConnection npgSqlConnection14 = new NpgsqlConnection(connectionString);
                            npgSqlConnection14.Open();
                            using (npgSqlConnection14)
                            {
                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression4, npgSqlConnection14);
                                // создаем параметры и добавляем их к команде
                                NpgsqlParameter Param1 = new NpgsqlParameter("@date", date_formate);
                                command.Parameters.Add(Param1);
                                NpgsqlParameter Param2 = new NpgsqlParameter("@characteristic", characteristic_date);
                                command.Parameters.Add(Param2);
                                NpgsqlParameter Param3 = new NpgsqlParameter("@fileReference", fileReference);
                                command.Parameters.Add(Param3);
                                NpgsqlParameter Param4 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                                command.Parameters.Add(Param4);
                                int number = command.ExecuteNonQuery();
                                //MessageBox.Show("Характеристика добавлена успешно!");

                            }
                            npgSqlConnection14.Close();
                        }
                    }

                    //Добавим гражданство в карточку гражданства
                    string SqlExpression101 = "INSERT INTO \"card-citizenship\" (\"pk_sitizenship\",\"pk_personal_card\") " +
                    "VALUES (@pk_citizenship,@pk_personal_card)";
                    NpgsqlConnection npgSqlConnection101 = new NpgsqlConnection(connectionString);
                    npgSqlConnection101.Open();
                    using (npgSqlConnection101)
                    {
                        NpgsqlCommand command = new NpgsqlCommand(SqlExpression101, npgSqlConnection101);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_citizenship", pk_citizenship);
                        command.Parameters.Add(Param1);
                        NpgsqlParameter Param2 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                        command.Parameters.Add(Param2);
                        int number = command.ExecuteNonQuery();
                        //MessageBox.Show("Гражданство добавлено успешно!");
                    }
                    npgSqlConnection101.Close();
                    MessageBox.Show("Сотрудник добавлен успешно!");


                }
                //Если форма открыта для редактирования добавляем нового сотрудника
                if (flag==1)
                {
                    string sqlExpression = "UPDATE \"PersonalCard\" SET \"pk_marital_status\"=@pk_marital_status, \"pk_character_work\"=@pk_character_work," +
                        "\"surname\"=@surname,\"name\"=@name,\"otchestvo\"=@otchestvo,\"birthday\"=@birthday,\"Characteristic\"=@Characteristic," +
                        "\"INN\"=@INN,\"SSN\"=@SSN,\"Serial_number\"=@Serial_number,\"Passport_date\"=@Passport_date,\"Vidan\"=@Vidan,\"Home_date\"=@Home_date," +
                        "\"Propiska\"=@Propiska,\"Fact_address\"=@Fact_address,\"Phone\"=@Phone,\"pk_military_rank\"=@pk_military_rank," +
                        "\"pk_category_military\"=@pk_category_military,\"pk_stock_category\"=@pk_stock_category,\"Birth_place\"=@Birth_place," +
                        "\"Creation_date\"=@Creation_date,\"Gender\"=@Gender,\"Military_profile\"=@Military_profile,\"Military_code\"=@Military_code," +
                        "\"Military_name\"=@Military_name,\"Military_status\"=@Military_status,\"Military_cancel\"=@Military_cancel,\"Work_kind\"=@Work_kind," +
                        "\"Index_fact\"=@Index_fact,\"Index_real\"=@Index_real WHERE \"pk_personal_card\"=@pk_personal_card";
                    using (npgSqlConnection)
                    {
                        NpgsqlCommand command = new NpgsqlCommand(sqlExpression, npgSqlConnection);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_marital_status", pk_marital_status);
                        command.Parameters.Add(Param1);
                        NpgsqlParameter Param2 = new NpgsqlParameter("@pk_character_work", pk_character_work);
                        command.Parameters.Add(Param2);
                        NpgsqlParameter Param3 = new NpgsqlParameter("@surname", surname);
                        command.Parameters.Add(Param3);
                        NpgsqlParameter Param4 = new NpgsqlParameter("@name", name);
                        command.Parameters.Add(Param4);
                        NpgsqlParameter Param5 = new NpgsqlParameter("@otchestvo", otchestvo);
                        command.Parameters.Add(Param5);
                        NpgsqlParameter Param6 = new NpgsqlParameter("@birthday", Date_birth);
                        command.Parameters.Add(Param6);
                        NpgsqlParameter Param7 = new NpgsqlParameter("@Characteristic", Characteristic);
                        command.Parameters.Add(Param7);
                        NpgsqlParameter Param8 = new NpgsqlParameter("@INN", INN);
                        command.Parameters.Add(Param8);
                        NpgsqlParameter Param9 = new NpgsqlParameter("@SSN", SNN);
                        command.Parameters.Add(Param9);
                        NpgsqlParameter Param10 = new NpgsqlParameter("@Serial_number", Serial_number);
                        command.Parameters.Add(Param10);
                        NpgsqlParameter Param11 = new NpgsqlParameter("@Passport_date", Passport_date);
                        command.Parameters.Add(Param11);
                        NpgsqlParameter Param12 = new NpgsqlParameter("@Vidan", Vidan);
                        command.Parameters.Add(Param12);
                        NpgsqlParameter Param13 = new NpgsqlParameter("@Home_date", Home_date);
                        command.Parameters.Add(Param13);
                        NpgsqlParameter Param14 = new NpgsqlParameter("@Propiska", Propiska);
                        command.Parameters.Add(Param14);
                        NpgsqlParameter Param15 = new NpgsqlParameter("@Fact_address", Fact_address);
                        command.Parameters.Add(Param15);
                        NpgsqlParameter Param16 = new NpgsqlParameter("@Phone", Phone);
                        command.Parameters.Add(Param16);
                        NpgsqlParameter Param17 = new NpgsqlParameter("@pk_military_rank", pk_military_rank);
                        command.Parameters.Add(Param17);
                        NpgsqlParameter Param18 = new NpgsqlParameter("@pk_category_military", pk_category_military);
                        command.Parameters.Add(Param18);
                        NpgsqlParameter Param19 = new NpgsqlParameter("@pk_stock_category", pk_stock_category);
                        command.Parameters.Add(Param19);
                        NpgsqlParameter Param20 = new NpgsqlParameter("@Birth_place", Birth_place);
                        command.Parameters.Add(Param20);
                        NpgsqlParameter Param21 = new NpgsqlParameter("@Creation_date", Creation_date);
                        command.Parameters.Add(Param21);
                        NpgsqlParameter Param22 = new NpgsqlParameter("@Gender", Gender);
                        command.Parameters.Add(Param22);
                        NpgsqlParameter Param23 = new NpgsqlParameter("@Military_profile", Military_profile);
                        command.Parameters.Add(Param23);
                        NpgsqlParameter Param24 = new NpgsqlParameter("@Military_code", Military_code);
                        command.Parameters.Add(Param24);
                        NpgsqlParameter Param25 = new NpgsqlParameter("@Military_name", Military_name);
                        command.Parameters.Add(Param25);
                        NpgsqlParameter Param26 = new NpgsqlParameter("@Military_status", Military_status);
                        command.Parameters.Add(Param26);
                        NpgsqlParameter Param27 = new NpgsqlParameter("@Military_cancel", Military_cancel);
                        command.Parameters.Add(Param27);
                        NpgsqlParameter Param28 = new NpgsqlParameter("@Work_kind", Work_kind);
                        command.Parameters.Add(Param28);
                        NpgsqlParameter Param29 = new NpgsqlParameter("@Index_fact", Index_fact);
                        command.Parameters.Add(Param29);
                        NpgsqlParameter Param30 = new NpgsqlParameter("@Index_real", Index_real);
                        command.Parameters.Add(Param30);
                        NpgsqlParameter Param31 = new NpgsqlParameter("@pk_personal_card",pk_personal_card );
                        command.Parameters.Add(Param31);


                        int number = command.ExecuteNonQuery();

                       


                    }
                    string sqlExpression900 = "DELETE FROM \"lang-card\" WHERE \"pk_personal_card\" = @pk_personal_card";
                    NpgsqlConnection npgSqlConnection900 = new NpgsqlConnection(connectionString);
                    npgSqlConnection900.Open();
                    using (npgSqlConnection900)
                    {
                        NpgsqlCommand command900 = new NpgsqlCommand(sqlExpression900, npgSqlConnection900);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                        command900.Parameters.Add(Param1);
                        int number = command900.ExecuteNonQuery();
                    }
                    npgSqlConnection900.Close();
                    //Проверка, были ли добавлены языки. Если да, то добавляем информацию в базу
                    if (dataGridView3.Rows.Count != 0)
                        {
                        foreach (DataGridViewRow Row in dataGridView3.Rows)
                        {
                            string language = Row.Cells[0].Value.ToString();
                            string language_degree = Row.Cells[1].Value.ToString();
                            int pk_language = -1;
                            int pk_degree_language = -1;
                            string SqlExpression = "SELECT * FROM public.\"Language\" WHERE \"Name\"=@language";
                            NpgsqlConnection npgSqlConnection7 = new NpgsqlConnection(connectionString);
                            npgSqlConnection7.Open();
                            using (npgSqlConnection7)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection7);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@language", language);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader7 = command.ExecuteReader();
                                if (reader7.HasRows) // если есть данные
                                {
                                    while (reader7.Read()) // построчно считываем данные
                                    {
                                        object language1 = reader7.GetValue(0);
                                        pk_language = Convert.ToInt32(language1);

                                    }
                                }


                            }
                            npgSqlConnection7.Close();

                            string SqlExpression1 = "SELECT * FROM public.\"DegreeLanguage\" WHERE \"Name\"=@language_degree";
                            NpgsqlConnection npgSqlConnection8 = new NpgsqlConnection(connectionString);
                            npgSqlConnection8.Open();
                            using (npgSqlConnection8)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression1, npgSqlConnection8);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@language_degree", language_degree);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader8 = command.ExecuteReader();
                                if (reader8.HasRows) // если есть данные
                                {
                                    while (reader8.Read()) // построчно считываем данные
                                    {
                                        object degree_language1 = reader8.GetValue(0);
                                        pk_degree_language = Convert.ToInt32(degree_language1);

                                    }
                                }


                            }
                            npgSqlConnection8.Close();

                            string SqlExpression2 = "INSERT INTO \"lang-card\" (\"pk_language\",\"pk_personal_card\",\"pk_degree_language\") " +
                                "VALUES (@pk_language,@pk_personal_card,@pk_degree_language)";
                            NpgsqlConnection npgSqlConnection9 = new NpgsqlConnection(connectionString);
                            npgSqlConnection9.Open();
                            using (npgSqlConnection9)
                            {
                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression2, npgSqlConnection9);
                                // создаем параметры и добавляем их к команде
                                NpgsqlParameter Param1 = new NpgsqlParameter("@pk_language", pk_language);
                                command.Parameters.Add(Param1);
                                NpgsqlParameter Param2 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                                command.Parameters.Add(Param2);
                                NpgsqlParameter Param3 = new NpgsqlParameter("@pk_degree_language", pk_degree_language);
                                command.Parameters.Add(Param3);
                                int number = command.ExecuteNonQuery();
                                //MessageBox.Show("Знание языков добавлено успешно!");

                            }
                            npgSqlConnection9.Close();
                        }
                        string sqlExpression901 = "DELETE FROM \"card-education\" WHERE \"pk_personal_card\" = @pk_personal_card";
                        NpgsqlConnection npgSqlConnection901 = new NpgsqlConnection(connectionString);
                        npgSqlConnection901.Open();
                        using (npgSqlConnection901)
                        {
                            NpgsqlCommand command901 = new NpgsqlCommand(sqlExpression901, npgSqlConnection901);
                            // создаем параметры и добавляем их к команде
                            NpgsqlParameter Param1 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                            command901.Parameters.Add(Param1);
                            int number = command901.ExecuteNonQuery();
                        }
                        npgSqlConnection901.Close();
                    }
                    //Проверка, были ли добавлены образования. Если да, то добавляем информацию в базу
                    if (dataGridView4.Rows.Count != 0)
                    {
                        foreach (DataGridViewRow Row in dataGridView4.Rows)
                        {
                            string document_name = Row.Cells[0].Value.ToString();
                            string education = Row.Cells[1].Value.ToString();
                            string institution = Row.Cells[2].Value.ToString();
                            string specialty = Row.Cells[3].Value.ToString();
                            string serial_number = Row.Cells[4].Value.ToString();
                            int Year = Convert.ToInt32(Row.Cells[5].Value);
                            int pk_education = -1;
                            int pk_specialty = -1;
                            int pk_nstitution = -1;
                            string SqlExpression = "SELECT * FROM public.\"Education\" WHERE \"Name\"=@education";
                            NpgsqlConnection npgSqlConnection10 = new NpgsqlConnection(connectionString);
                            npgSqlConnection10.Open();
                            using (npgSqlConnection10)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression, npgSqlConnection10);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@education", education);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader10 = command.ExecuteReader();
                                if (reader10.HasRows) // если есть данные
                                {
                                    while (reader10.Read()) // построчно считываем данные
                                    {
                                        object education1 = reader10.GetValue(0);
                                        pk_education = Convert.ToInt32(education1);

                                    }
                                }


                            }
                            npgSqlConnection10.Close();
                            string SqlExpression1 = "SELECT * FROM public.\"Specialty\" WHERE \"Name\"=@specialty";
                            NpgsqlConnection npgSqlConnection11 = new NpgsqlConnection(connectionString);
                            npgSqlConnection11.Open();
                            using (npgSqlConnection11)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression1, npgSqlConnection11);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@specialty", specialty);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader11 = command.ExecuteReader();
                                if (reader11.HasRows) // если есть данные
                                {
                                    while (reader11.Read()) // построчно считываем данные
                                    {
                                        object speciality1 = reader11.GetValue(0);
                                        pk_specialty = Convert.ToInt32(speciality1);

                                    }
                                }


                            }
                            npgSqlConnection11.Close();
                            string SqlExpression2 = "SELECT * FROM public.\"Institution\" WHERE \"Name\"=@institution";
                            NpgsqlConnection npgSqlConnection12 = new NpgsqlConnection(connectionString);
                            npgSqlConnection12.Open();
                            using (npgSqlConnection12)
                            {

                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression2, npgSqlConnection12);
                                // создаем параметр для имени
                                NpgsqlParameter CWParam = new NpgsqlParameter("@institution", institution);
                                command.Parameters.Add(CWParam);
                                NpgsqlDataReader reader12 = command.ExecuteReader();
                                if (reader12.HasRows) // если есть данные
                                {
                                    while (reader12.Read()) // построчно считываем данные
                                    {
                                        object institution1 = reader12.GetValue(0);
                                        pk_nstitution = Convert.ToInt32(institution1);

                                    }
                                }


                            }
                            npgSqlConnection12.Close();
                            string SqlExpression3 = "INSERT INTO \"card-education\" (\"pk_education\",\"pk_personal_card\",\"pk_specialty\"," +
                                "\"pk_nstitution\",\"document_name\",\"serial_number\",\"Year\") " +
                                "VALUES (@pk_education,@pk_personal_card,@pk_specialty,@pk_nstitution,@document_name,@serial_number,@Year)";
                            NpgsqlConnection npgSqlConnection13 = new NpgsqlConnection(connectionString);
                            npgSqlConnection13.Open();
                            using (npgSqlConnection13)
                            {
                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression3, npgSqlConnection13);
                                // создаем параметры и добавляем их к команде
                                NpgsqlParameter Param1 = new NpgsqlParameter("@pk_education", pk_education);
                                command.Parameters.Add(Param1);
                                NpgsqlParameter Param2 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                                command.Parameters.Add(Param2);
                                NpgsqlParameter Param3 = new NpgsqlParameter("@pk_specialty", pk_specialty);
                                command.Parameters.Add(Param3);
                                NpgsqlParameter Param4 = new NpgsqlParameter("@pk_nstitution", pk_nstitution);
                                command.Parameters.Add(Param4);
                                NpgsqlParameter Param5 = new NpgsqlParameter("@document_name", document_name);
                                command.Parameters.Add(Param5);
                                NpgsqlParameter Param6 = new NpgsqlParameter("@serial_number", serial_number);
                                command.Parameters.Add(Param6);
                                NpgsqlParameter Param7 = new NpgsqlParameter("@Year", Year);
                                command.Parameters.Add(Param7);
                                int number = command.ExecuteNonQuery();
                                //MessageBox.Show("Образование добавлено успешно!");

                            }
                            npgSqlConnection13.Close();
                        }
                    }
                    string sqlExpression902 = "DELETE FROM \"Characteristic\" WHERE \"pk_personal_card\" = @pk_personal_card";
                    NpgsqlConnection npgSqlConnection902 = new NpgsqlConnection(connectionString);
                    npgSqlConnection902.Open();
                    using (npgSqlConnection902)
                    {
                        NpgsqlCommand command902 = new NpgsqlCommand(sqlExpression902, npgSqlConnection902);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                        command902.Parameters.Add(Param1);
                        int number = command902.ExecuteNonQuery();
                    }
                    npgSqlConnection902.Close();
                    //Проверка, были ли добавлены характеристики. Если да, то добавляем информацию в базу
                    if (dataGridView5.Rows.Count != 0)
                    {
                        foreach (DataGridViewRow Row in dataGridView5.Rows)
                        {
                            string date = Row.Cells[0].Value.ToString();
                            DateTime date_formate = Convert.ToDateTime(date);
                            string fileReference = Row.Cells[1].Value.ToString();
                            string characteristic_date;
                            FileStream stream = new FileStream(Row.Cells[1].Value.ToString(), FileMode.Open);
                            StreamReader reader = new StreamReader(stream);
                            characteristic_date = (reader.ReadToEnd());
                            richTextBox1.Text = (reader.ReadToEnd());
                            reader.Close();
                            string SqlExpression4 = "INSERT INTO \"Characteristic\" (\"date\",\"characteristic\",\"fileReference\"," +
                                "\"pk_personal_card\") " +
                                "VALUES (@date,@characteristic,@fileReference,@pk_personal_card)";
                            NpgsqlConnection npgSqlConnection14 = new NpgsqlConnection(connectionString);
                            npgSqlConnection14.Open();
                            using (npgSqlConnection14)
                            {
                                NpgsqlCommand command = new NpgsqlCommand(SqlExpression4, npgSqlConnection14);
                                // создаем параметры и добавляем их к команде
                                NpgsqlParameter Param1 = new NpgsqlParameter("@date", date_formate);
                                command.Parameters.Add(Param1);
                                NpgsqlParameter Param2 = new NpgsqlParameter("@characteristic", characteristic_date);
                                command.Parameters.Add(Param2);
                                NpgsqlParameter Param3 = new NpgsqlParameter("@fileReference", fileReference);
                                command.Parameters.Add(Param3);
                                NpgsqlParameter Param4 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                                command.Parameters.Add(Param4);
                                int number = command.ExecuteNonQuery();
                                //MessageBox.Show("Характеристика добавлена успешно!");

                            }
                            npgSqlConnection14.Close();
                        }
                    }
                    string sqlExpression903 = "DELETE FROM \"card-citizenship\" WHERE \"pk_personal_card\" = @pk_personal_card";
                    NpgsqlConnection npgSqlConnection903 = new NpgsqlConnection(connectionString);
                    npgSqlConnection903.Open();
                    using (npgSqlConnection903)
                    {
                        NpgsqlCommand command903 = new NpgsqlCommand(sqlExpression903, npgSqlConnection903);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                        command903.Parameters.Add(Param1);
                        int number = command903.ExecuteNonQuery();
                    }
                    npgSqlConnection903.Close();
                    //Добавим гражданство в карточку гражданства
                    string SqlExpression101 = "INSERT INTO \"card-citizenship\" (\"pk_sitizenship\",\"pk_personal_card\") " +
                    "VALUES (@pk_citizenship,@pk_personal_card)";
                    NpgsqlConnection npgSqlConnection101 = new NpgsqlConnection(connectionString);
                    npgSqlConnection101.Open();
                    using (npgSqlConnection101)
                    {
                        NpgsqlCommand command = new NpgsqlCommand(SqlExpression101, npgSqlConnection101);
                        // создаем параметры и добавляем их к команде
                        NpgsqlParameter Param1 = new NpgsqlParameter("@pk_citizenship", pk_citizenship);
                        command.Parameters.Add(Param1);
                        NpgsqlParameter Param2 = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                        command.Parameters.Add(Param2);
                        int number = command.ExecuteNonQuery();
                        //MessageBox.Show("Гражданство добавлено успешно!");
                    }
                    npgSqlConnection101.Close();
                    MessageBox.Show("Редактирование сотрудника успешно!");

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
			String connectionString = "Server=hrd.cx7kyl76gv42.us-east-2.rds.amazonaws.com;User Id=" + Connection.login + ";Password=" + Connection.password + ";Database=HRD;";
			NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            int pk_marital_status = -1;
            int pk_character_work = -1;
            int pk_military_rank = -1;
            int pk_category_military = -1;
            int pk_stock_category = -1;
            int pk_citizenship = -1;
            if (flag == 0)
            {
                tabPage9.Parent = null; //При добавлении скрываем таблицу должностей
                panel2.Visible = false; //При добавлении еще нет текущей должности скрываем панель с ней
            }
            //Если форма открыта для редактирования считываем из базы данных информацию о сотруднике
            if (flag == 1)
            {
                
                try
                {
                    //Вот тут я запросом считываю из базы личную карточку
                    string sqlExpression20 = "SELECT * FROM public.\"PersonalCard\" WHERE \"pk_personal_card\"=@pk_personal_card";
                    npgSqlConnection.Open();
                    NpgsqlCommand command20 = new NpgsqlCommand(sqlExpression20, npgSqlConnection);
                    NpgsqlParameter CWParam = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                    command20.Parameters.Add(CWParam);
                    NpgsqlDataReader reader20 = command20.ExecuteReader();
                    if (reader20.HasRows) // если есть данные
                    {
                        while (reader20.Read()) // построчно считываем данные
                        {
                            object marital_status = reader20.GetValue(1);
                            pk_marital_status = Convert.ToInt32(marital_status);
                            object character_work = reader20.GetValue(2);
                            pk_character_work = Convert.ToInt32(character_work);
                            object surname = reader20.GetValue(3);
                            richTextBox14.Text = surname.ToString();
                            object name = reader20.GetValue(4);
                            richTextBox15.Text = name.ToString();
                            object otchestvo = reader20.GetValue(5);
                            richTextBox13.Text = otchestvo.ToString();
                            object birthday_date = reader20.GetValue(6);
                            dateTimePicker6.Value = Convert.ToDateTime(birthday_date);
                            object INN = reader20.GetValue(8);
                            richTextBox4.Text = INN.ToString();
                            object SSN = reader20.GetValue(9);
                            richTextBox3.Text = SSN.ToString();
                            object Serial_number = reader20.GetValue(11);
                            richTextBox6.Text = Serial_number.ToString();
                            object Passport_date = reader20.GetValue(12);
                            dateTimePicker4.Value = Convert.ToDateTime(Passport_date);
                            object Vidan = reader20.GetValue(13);
                            richTextBox5.Text = Vidan.ToString();
                            object Home_date = reader20.GetValue(14);
                            dateTimePicker1.Value = Convert.ToDateTime(Home_date);
                            object Propiska = reader20.GetValue(15);
                            richTextBox9.Text = Propiska.ToString();
                            object Fact_address = reader20.GetValue(16);
                            richTextBox11.Text = Fact_address.ToString();
                            object Phone = reader20.GetValue(17);
                            richTextBox2.Text = Phone.ToString();
                            object military_rank = reader20.GetValue(18);
                            pk_military_rank = Convert.ToInt32(military_rank);
                            object category_military = reader20.GetValue(19);
                            pk_category_military = Convert.ToInt32(category_military);
                            object stock_category = reader20.GetValue(20);
                            pk_stock_category = Convert.ToInt32(stock_category);
                            object Birth_place = reader20.GetValue(21);
                            richTextBox7.Text = Birth_place.ToString();
                            object Creation_date = reader20.GetValue(22);
                            dateTimePicker2.Value = Convert.ToDateTime(Creation_date);
                            object Gender = reader20.GetValue(23);
                            if (Gender.ToString() == "М") radioButton3.Checked = true;
                            if (Gender.ToString() == "Ж") radioButton4.Checked = true;
                            object Military_profile = reader20.GetValue(24);
                            richTextBox16.Text = Military_profile.ToString();
                            object Military_code = reader20.GetValue(25);
                            richTextBox17.Text = Military_code.ToString();
                            object Military_name = reader20.GetValue(26);
                            richTextBox18.Text = Military_name.ToString();
                            object Military_status = reader20.GetValue(27);
                            comboBox11.Text = Military_status.ToString();
                            object Military_cancel = reader20.GetValue(28);
                            richTextBox19.Text = Military_cancel.ToString();
                            object Work_kind = reader20.GetValue(29);
                            richTextBox12.Text = Work_kind.ToString();
                            object Index_fact = reader20.GetValue(30);
                            richTextBox10.Text = Index_fact.ToString();
                            object Index_real = reader20.GetValue(31);
                            richTextBox8.Text = Index_real.ToString();

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
                    //Вот тут я запросом считываю из базы языки 
                    string sqlExpression21 = "SELECT \"Language\".\"Name\",\"DegreeLanguage\".\"Name\"" +
                        "FROM \"Language\",\"DegreeLanguage\",\"lang-card\"" +
                        "WHERE \"lang-card\".\"pk_personal_card\"=@pk_personal_card and \"lang-card\".\"pk_language\"=\"Language\".\"pk_language\"" +
                        "and \"lang-card\".\"pk_degree_language\"=\"DegreeLanguage\".\"pk_degree_language\""; 
                    npgSqlConnection.Open();
                    // MessageBox.Show("Подключение открыто!!");
                    NpgsqlCommand command21 = new NpgsqlCommand(sqlExpression21, npgSqlConnection);
                    NpgsqlParameter CWParam = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                    command21.Parameters.Add(CWParam);
                    NpgsqlDataReader reader21 = command21.ExecuteReader();
                    if (reader21.HasRows) // если есть данные
                    {
                        while (reader21.Read()) // построчно считываем данные
                        {
                            object Lang_name = reader21.GetValue(0);
                            object Degree_name = reader21.GetValue(1);
                            dataGridView3.Rows.Add();
                            int kol = dataGridView3.Rows.Count - 1;
                            dataGridView3.Rows[kol].Cells[0].Value = Lang_name;
                            dataGridView3.Rows[kol].Cells[1].Value = Degree_name;

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
                    //Вот тут я запросом считываю из базы карточку-гражданство сотрудника
                    string sqlExpression22 = "SELECT * FROM public.\"card-citizenship\" WHERE \"pk_personal_card\"=@pk_personal_card";
                    npgSqlConnection.Open();
                    // MessageBox.Show("Подключение открыто!!");
                    NpgsqlCommand command22 = new NpgsqlCommand(sqlExpression22, npgSqlConnection);
                    NpgsqlParameter CWParam = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                    command22.Parameters.Add(CWParam);
                    NpgsqlDataReader reader22 = command22.ExecuteReader();
                    if (reader22.HasRows) // если есть данные
                    {
                        while (reader22.Read()) // построчно считываем данные
                        {
                            object pk = reader22.GetValue(1);
                            pk_citizenship = Convert.ToInt32(pk);
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
                    //Вот тут я запросом считываю из базы образования 
                    string sqlExpression23 = "SELECT \"card-education\".\"document_name\",\"Education\".\"Name\",\"Institution\".\"Name\"" +
                        ",\"Specialty\".\"Name\",\"card-education\".\"serial_number\",\"card-education\".\"Year\"" +
                        "FROM \"card-education\",\"Education\",\"Institution\",\"Specialty\"" +
                        "WHERE \"card-education\".\"pk_personal_card\"=@pk_personal_card and \"card-education\".\"pk_education\"=\"Education\".\"pk_education\"" +
                        "and \"card-education\".\"pk_nstitution\"=\"Institution\".\"pk_nstitution\" " +
                        "and \"card-education\".\"pk_specialty\"=\"Specialty\".\"pk_specialty\"";
                    npgSqlConnection.Open();
                    // MessageBox.Show("Подключение открыто!!");
                    NpgsqlCommand command23 = new NpgsqlCommand(sqlExpression23, npgSqlConnection);
                    NpgsqlParameter CWParam = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                    command23.Parameters.Add(CWParam);
                    NpgsqlDataReader reader23 = command23.ExecuteReader();
                    if (reader23.HasRows) // если есть данные
                    {
                        while (reader23.Read()) // построчно считываем данные
                        {
                            object document_name = reader23.GetValue(0);
                            object education_name = reader23.GetValue(1);
                            object institution_name = reader23.GetValue(2);
                            object specialty_name = reader23.GetValue(3);
                            object serial_number = reader23.GetValue(4);
                            object Year = reader23.GetValue(5);
                            dataGridView4.Rows.Add();
                            int kol = dataGridView4.Rows.Count - 1;
                            dataGridView4.Rows[kol].Cells[0].Value = document_name;
                            dataGridView4.Rows[kol].Cells[1].Value = education_name;
                            dataGridView4.Rows[kol].Cells[2].Value = institution_name;
                            dataGridView4.Rows[kol].Cells[3].Value = specialty_name;
                            dataGridView4.Rows[kol].Cells[4].Value = serial_number;
                            dataGridView4.Rows[kol].Cells[5].Value = Year;
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
                    //Вот тут я запросом считываю из базы характеристики
                    string sqlExpression24 = "SELECT * FROM public.\"Characteristic\" WHERE \"pk_personal_card\"=@pk_personal_card";
                    npgSqlConnection.Open();
                    // MessageBox.Show("Подключение открыто!!");
                    NpgsqlCommand command24 = new NpgsqlCommand(sqlExpression24, npgSqlConnection);
                    NpgsqlParameter CWParam = new NpgsqlParameter("@pk_personal_card", pk_personal_card);
                    command24.Parameters.Add(CWParam);
                    NpgsqlDataReader reader24 = command24.ExecuteReader();
                    if (reader24.HasRows) // если есть данные
                    {
                        while (reader24.Read()) // построчно считываем данные
                        {
                            object id_characteristic = reader24.GetValue(0);
                            object date = reader24.GetValue(1);
                            object Characteristic= reader24.GetValue(2);
                            object fileReference = reader24.GetValue(3);
                            string path = @"C:\KADRY_CHARACTERISTIC";
                            DirectoryInfo dirInfo = new DirectoryInfo(path);
                            if (!dirInfo.Exists)
                            {
                                dirInfo.Create();
                                
                            }
                            System.IO.File.WriteAllText("C:\\KADRY_CHARACTERISTIC\\Characteristic" + id_characteristic.ToString() + ".txt", Characteristic.ToString());
                            dataGridView5.Rows.Add();
                            int kol = dataGridView5.Rows.Count - 1;
                            dataGridView5.Rows[kol].Cells[0].Value = date;
                            dataGridView5.Rows[kol].Cells[1].Value = "C:\\KADRY_CHARACTERISTIC\\Characteristic" + id_characteristic.ToString() + ".txt";
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
                        if (flag == 1)
                        {
                            object Id = reader.GetValue(0);
                            if (Convert.ToInt32(Id) == pk_citizenship)
                            {
                                object Name1 = reader.GetValue(1);
                                comboBox9.Text = Name1.ToString();
                            }
                        }
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
                        if (flag == 1)
                        {
                            object Id = reader1.GetValue(0);
                            if (Convert.ToInt32(Id) == pk_marital_status)
                            {
                                object Name1 = reader1.GetValue(1);
                                comboBox7.Text=Name1.ToString();
                            }
                        }
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
                        if (flag == 1)
                        {
                            object Id = reader2.GetValue(0);
                            if (Convert.ToInt32(Id) == pk_character_work)
                            {
                                object Name1 = reader2.GetValue(1);
                                comboBox4.Text = Name1.ToString();
                            }
                        }
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

            try
            {
                //Вот тут я запросом считываю из базы категорию запаса в comboBox
                string sqlExpression3 = "SELECT * FROM public.\"StockCategory\"";
                npgSqlConnection.Open();
                NpgsqlCommand command3 = new NpgsqlCommand(sqlExpression3, npgSqlConnection);
                NpgsqlDataReader reader3 = command3.ExecuteReader();
                if (reader3.HasRows) // если есть данные
                {
                    while (reader3.Read()) // построчно считываем данные
                    {
                        if (flag == 1)
                        {
                            object Id = reader3.GetValue(0);
                            if (Convert.ToInt32(Id) == pk_stock_category)
                            {
                                object Name1 = reader3.GetValue(1);
                                comboBox8.Text = Name1.ToString();
                            }
                        }
                        object Name = reader3.GetValue(1);
                        comboBox8.Items.Add(Name);
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
                //Вот тут я запросом считываю из базы категорию годности в comboBox
                string sqlExpression4 = "SELECT * FROM public.\"CategoryMilitary\"";
                npgSqlConnection.Open();
                NpgsqlCommand command4 = new NpgsqlCommand(sqlExpression4, npgSqlConnection);
                NpgsqlDataReader reader4 = command4.ExecuteReader();
                if (reader4.HasRows) // если есть данные
                {
                    while (reader4.Read()) // построчно считываем данные
                    {
                        if (flag == 1)
                        {
                            object Id = reader4.GetValue(0);
                            if (Convert.ToInt32(Id) == pk_category_military)
                            {
                                object Name1 = reader4.GetValue(1);
                                comboBox5.Text = Name1.ToString();
                            }
                        }
                        object Name = reader4.GetValue(1);
                        comboBox5.Items.Add(Name);
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
                //Вот тут я запросом считываю из базы воинское звание в comboBox
                string sqlExpression5 = "SELECT * FROM public.\"MilitaryRank\"";
                npgSqlConnection.Open();
                NpgsqlCommand command5 = new NpgsqlCommand(sqlExpression5, npgSqlConnection);
                NpgsqlDataReader reader5 = command5.ExecuteReader();
                if (reader5.HasRows) // если есть данные
                {
                    while (reader5.Read()) // построчно считываем данные
                    {
                        if (flag == 1)
                        {
                            object Id = reader5.GetValue(0);
                            if (Convert.ToInt32(Id) == pk_military_rank)
                            {
                                object Name1 = reader5.GetValue(1);
                                comboBox6.Text = Name1.ToString();
                            }
                        }
                        object Name = reader5.GetValue(1);
                        comboBox6.Items.Add(Name);
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

            comboBox11.Items.Add("Состоит");
            comboBox11.Items.Add("Не состоит");
            

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

        //Обработка нажатия сохранить изменения в графе характеристика
        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Текстовый документ (*.txt)|*.txt|Все файлы (*.*)|*.*";
            string FileName="";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName);
                FileName = saveFileDialog.FileName;
                streamWriter.WriteLine(richTextBox1.Text+"\n");
                streamWriter.Close();

            }

            if (flagnewfile!=-1)
            {
                dataGridView5.Rows[flagnewfile].Cells[0].Value = dateTimePicker2.Value;
                dataGridView5.Rows[flagnewfile].Cells[1].Value = FileName;
            }
            if (flagnewfile == -1)
            {
                
                dataGridView5.Rows.Add();
                int kol = dataGridView5.Rows.Count - 1;
                flagnewfile = kol;
                dataGridView5.Rows[kol].Cells[0].Value = dateTimePicker2.Value;
                dataGridView5.Rows[kol].Cells[1].Value = FileName;
            }
        }

        private void richTextBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form ifrm = Application.OpenForms[Application.OpenForms.Count - 1];
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

        //Ограничение на ввод индекса в фактическом адресе
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
            if (richTextBox3.Text.Length > 13)
            {
                e.Handled = true;
            }
          
        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            richTextBox7.Text = listBox1.SelectedItem.ToString();
        }

        //Открытие формы знание иностранного языка
        private void button3_Click(object sender, EventArgs e)
        {
            Form5 f5 = new Form5();
            f5.Owner = this;
            f5.Show();
            this.Hide();
        }
        //Удаление языка из таблицы
        private void button2_Click(object sender, EventArgs e)
        {
            if ((dataGridView3.Rows.Count == 0))
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            if (dataGridView3.SelectedRows.Count < 0)
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            else
            {
                int a = dataGridView3.CurrentRow.Index;
                dataGridView3.Rows.Remove(dataGridView3.Rows[a]);
            }
        }

        //Открытие формы образование
        private void button11_Click(object sender, EventArgs e)
        {
            Form6 f6 = new Form6();
            f6.Owner = this;
            f6.Show();
            this.Hide();
        }
        //Удаление образования из таблицы
        private void button7_Click(object sender, EventArgs e)
        {
            if ((dataGridView4.Rows.Count == 0))
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            if (dataGridView4.SelectedRows.Count < 0)
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            else
            {
                int a = dataGridView4.CurrentRow.Index;
                dataGridView4.Rows.Remove(dataGridView4.Rows[a]);
            }
        }

        //Обработка нажатия добавить файл с компьютера во вкладке характеристика
        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Текстовый документ (*.txt)|*.txt|Все файлы (*.*)|*.*";
            string FileName = "";
            int flagcheckfile = 0;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                flagcheckfile = 1;
                StreamReader streamReader = new StreamReader(openFileDialog.FileName);
                FileName = openFileDialog.FileName;
                richTextBox1.Text= (streamReader.ReadToEnd());
                streamReader.Close();

            }
            if (flagcheckfile == 1)
            {
                dataGridView5.Rows.Add();
                int kol = dataGridView5.Rows.Count - 1;
                flagnewfile = kol;
                dataGridView5.Rows[kol].Cells[0].Value = dateTimePicker2.Value;
                dataGridView5.Rows[kol].Cells[1].Value = FileName;
            }

        }

        //Обработка нажатия открыть файл из таблицы во вкладке характеристика
        private void button8_Click(object sender, EventArgs e)
        {
            if ((dataGridView5.Rows.Count == 0))
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            if (dataGridView5.SelectedRows.Count < 0)
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            else
            {
                flagnewfile = dataGridView5.CurrentRow.Index;
                FileStream stream = new FileStream(dataGridView5.Rows[flagnewfile].Cells[1].Value.ToString(), FileMode.Open);
                StreamReader reader = new StreamReader(stream);
                richTextBox1.Text = reader.ReadToEnd();
                stream.Close();

            }
        }

        //Обработка нажатия новая характеристика во вкладке характеристика
        private void button10_Click(object sender, EventArgs e)
        {
            flagnewfile = -1;
            richTextBox1.Text = "";
        }

        //Обработка нажатия Удалить файл из таблицы во вкладке характеристика
        private void button9_Click(object sender, EventArgs e)
        {
            if ((dataGridView5.Rows.Count == 0))
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            if (dataGridView5.SelectedRows.Count < 0)
            {
                MessageBox.Show("Не выбрана запись!");
                return;
            }
            else
            {
                int a = dataGridView5.CurrentRow.Index;
                if(a==flagnewfile)
                {
                    flagnewfile = -1;
                    richTextBox1.Text = "";
                }
                dataGridView5.Rows.Remove(dataGridView5.Rows[a]);
            }
        }

        
        private void richTextBox8_TextChanged(object sender, EventArgs e)
        {
            
            
        }

        //Защита индекса по паспорту для пользователей, кто любит ctrl-c ctrl-v 
        private void richTextBox8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {

                string Clipboard_data = Clipboard.GetText();
                foreach (var symbol in Clipboard_data)
                    if (!char.IsDigit(symbol))
                    {
                        MessageBox.Show("Индекс должен состоять только из цифр!");
                        break;
                    }

            }
        }

        
        private void richTextBox10_TextChanged(object sender, EventArgs e)
        {
            
        }

        //Защита индекса фактического адреса для пользователей, кто любит ctrl-c ctrl-v
        private void richTextBox10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {

                string Clipboard_data = Clipboard.GetText();
                foreach (var symbol in Clipboard_data)
                    if (!char.IsDigit(symbol))
                    {
                        MessageBox.Show("Индекс должен состоять только из цифр!");
                        break;
                    }

            }
        }

        //Защита ИНН для пользователей, кто любит ctrl-c ctrl-v
        private void richTextBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {

                string Clipboard_data = Clipboard.GetText();
                foreach (var symbol in Clipboard_data)
                    if (!char.IsDigit(symbol))
                    {
                        MessageBox.Show("ИНН должен состоять только из цифр!");
                        break;
                    }
                
            }
        }

        private void richTextBox9_TextChanged(object sender, EventArgs e)
        {
            var api = new SuggestClient(token);

            AutoCompleteStringCollection help = new AutoCompleteStringCollection();
            var response = api.SuggestAddress(richTextBox9.Text);

            listBox3.Items.Clear();
            for (int i = 0; i < response.suggestions.Count; i++)
            {
                listBox3.Items.Add(response.suggestions[i].value.ToString());
            }
        }

        private void richTextBox11_TextChanged_1(object sender, EventArgs e)
        {
            var api = new SuggestClient(token);

            AutoCompleteStringCollection help = new AutoCompleteStringCollection();
            var response = api.SuggestAddress(richTextBox11.Text);

            listBox4.Items.Clear();
            for (int i = 0; i < response.suggestions.Count; i++)
            {
                listBox4.Items.Add(response.suggestions[i].value.ToString());
            }
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            richTextBox9.Text = listBox3.SelectedItem.ToString();
        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            richTextBox11.Text = listBox4.SelectedItem.ToString();
        }
    }
}
