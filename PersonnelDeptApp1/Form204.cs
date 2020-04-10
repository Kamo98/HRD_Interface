using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using Npgsql;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace PersonnelDeptApp1
{
    public partial class Form204 : Form
    {
        Connection connectPSQL;
        NpgsqlConnection npgSqlConnection;
        public Form204()
        {
            InitializeComponent();
            connectPSQL = Connection.get_instance("postgres", "Ntcnbhjdfybt_01");
            npgSqlConnection = connectPSQL.get_connect();

            numericUpDown1.Value = DateTime.Now.Year;
            numericUpDown2.Value = DateTime.Now.Month;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Random rand = new Random();
            string pk_time_tracking = "";
            //ключ подразделения
            NpgsqlCommand com;
            com = new NpgsqlCommand("select \"pk_unit\" from \"Unit\" where \"Unit\".\"Name\" = '" + comboBox1.Text + "'", npgSqlConnection);
            string pk_unit = com.ExecuteScalar().ToString();

            //формируем дату
            string date = numericUpDown1.Value.ToString() + "-" + numericUpDown2.Value.ToString() + "-" + "01";

            bool flag_allow = false;
            //находим ключ табеля
            com = new NpgsqlCommand("SELECT \"pk_time_tracking\" FROM \"TimeTracking\" WHERE \"TimeTracking\".\"pk_unit\" = " + pk_unit + " AND \"TimeTracking\".\"from\" = '" + date + "'", npgSqlConnection);
            try
            {
                com.ExecuteScalar().ToString();
            }
            catch
            {
                //табеля не существует значит можно создавать
                flag_allow = true;
            }

            if (flag_allow)
            {
                //создание табеля

                MessageBox.Show("Создание табеля может занять некоторое время. Программа не зависла!");

                //дата по
                string date_to = numericUpDown1.Value.ToString() + "-" + numericUpDown2.Value.ToString() + "-" + DateTime.DaysInMonth(Convert.ToInt32(numericUpDown1.Value),Convert.ToInt32(numericUpDown2.Value));
                string nomer = rand.Next(10000000, 100000000).ToString();
                string date_sostav = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day;

                //вставка шапки табеля
                string queryIns = "insert into \"TimeTracking\" " +
                    "(\"nomer\", " +
                    "\"date_sostav\", " +
                    "\"from\"," +
                    " \"to\"," +
                    " \"pk_unit\")" +
                    " values ('" +
                    nomer + "', '" +
                    date_sostav + "', '" +
                    date + "', '" +
                    date_to + "', '" +
                    pk_unit + "') RETURNING \"pk_time_tracking\"";

                com = new NpgsqlCommand(queryIns, npgSqlConnection);
                NpgsqlDataReader reader = com.ExecuteReader();
                foreach (DbDataRecord rec in reader)
                {
                    pk_time_tracking = rec.GetInt32(0).ToString();
                }
                reader.Close();

                //поиск ключей сотрудников работающих в данном подразделении
                List<string> pk_pers_card = new List<string>();
                queryIns = "SELECT \"PeriodPosition\".\"pk_personal_card\" FROM \"PeriodPosition\",\"Position\" WHERE \"PeriodPosition\".\"pk_position\" = \"Position\".\"pk_position\" AND \"Position\".\"pk_unit\" = '" + pk_unit + "' AND \"PeriodPosition\".\"DateTo\" is null";
                com = new NpgsqlCommand(queryIns, npgSqlConnection);
                reader = com.ExecuteReader();
                if (reader.HasRows)
                {
                    foreach (DbDataRecord rec in reader)
                    {
                        pk_pers_card.Add(rec.GetInt32(0).ToString());
                    }
                }
                reader.Close();

                List<string> pk_string_tabel = new List<string>();
                //вставляем строки табеля
                for (int i = 0; i < pk_pers_card.Count; i++)
                {
                    queryIns = "insert into \"StringTimeTracking\" " +
                    "(\"pk_personal_card\", " +
                    " \"pk_time_tracking\")" +
                    " values ('" +
                    pk_pers_card[i] + "', '" +
                    pk_time_tracking + "') RETURNING \"pk_string_time_tracking\"";
                    com = new NpgsqlCommand(queryIns, npgSqlConnection);
                    reader = com.ExecuteReader();
                    foreach (DbDataRecord rec in reader)
                    {
                        pk_string_tabel.Add (rec.GetInt32(0).ToString());
                    }
                    reader.Close();
                }

                //определяем ключи шифров Я и В
                queryIns = "SELECT \"MarkTimeTracking\".\"pk_mark_time_tracking\" FROM \"MarkTimeTracking\" WHERE \"MarkTimeTracking\".\"ShortName\" = 'Я'";
                com = new NpgsqlCommand(queryIns, npgSqlConnection);
                string pk_shifr_YA = com.ExecuteScalar().ToString();
                queryIns = "SELECT \"MarkTimeTracking\".\"pk_mark_time_tracking\" FROM \"MarkTimeTracking\" WHERE \"MarkTimeTracking\".\"ShortName\" = 'В'";
                com = new NpgsqlCommand(queryIns, npgSqlConnection);
                string pk_shifr_V = com.ExecuteScalar().ToString();


                //вставляем факты в каждую строку
                for (int i = 0; i < pk_string_tabel.Count; i++)
                {
                    //сколько дней в месяце столько и вставок.
                    int daysInMonth =  DateTime.DaysInMonth(Convert.ToInt32(numericUpDown1.Value), Convert.ToInt32(numericUpDown2.Value));
                    for (int j = 0; j < daysInMonth; j++)
                    {
                        //формируем дату
                        string dt = numericUpDown1.Text + "-" + numericUpDown2.Text + "-" + (j + 1).ToString();

                        //определяем выходной или будни
                        DayOfWeek dayWeek = new DateTime(Convert.ToInt32(numericUpDown1.Value), Convert.ToInt32(numericUpDown2.Value), j + 1).DayOfWeek;
                        if (dayWeek == DayOfWeek.Saturday || dayWeek == DayOfWeek.Sunday)
                        {
                            queryIns = "insert into \"Fact\" " +
                                "(\"pk_string_time_tracking\", " +
                                " \"pk_mark_time_tracking\", " +
                                " \"data\"," +
                                " \"count_of_hours\")" +
                                " values ('" +
                                pk_string_tabel[i] + "', '" +
                                pk_shifr_V + "', '" +
                                dt + "', '" +
                             "0" + "')";
                        }
                        else
                        {
                            queryIns = "insert into \"Fact\" " +
                                "(\"pk_string_time_tracking\", " +
                                " \"pk_mark_time_tracking\", " +
                                " \"data\"," +
                                " \"count_of_hours\")" +
                                " values ('" +
                                pk_string_tabel[i] + "', '" +
                                pk_shifr_YA + "', '" +
                                dt + "', '" +
                             "8" + "')";
                        }
                        //вставляем явку
                        com = new NpgsqlCommand(queryIns, npgSqlConnection);
                        com.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Табель успешно создан!");
            }
            else
            {
                MessageBox.Show("Табель подразделения \"" + comboBox1.Text + "\" на " + date + " уже существует! Пожалуста воспользуйтесь поиском для его редактирования.");
                return;
            }
        }

        private void Form204_Load(object sender, EventArgs e)
        {
            AutoCompleteStringCollection listUnit = new AutoCompleteStringCollection();
            NpgsqlCommand com = new NpgsqlCommand("SELECT \"Name\" FROM \"Unit\"", npgSqlConnection);
            NpgsqlDataReader reader = com.ExecuteReader();

            if (reader.HasRows)
            {
                foreach (DbDataRecord rec in reader)
                {
                    listUnit.Add(rec.GetString(0));
                }

            }
            reader.Close();

            comboBox1.DataSource = listUnit;
            comboBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
            comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            comboBox1.AutoCompleteCustomSource = listUnit;
        }
    }
}
