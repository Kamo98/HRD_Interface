using System;
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

namespace PersonnelDeptApp1
{
    public partial class Form3 : Form
    {
        Connection connectPSQL;
        NpgsqlConnection npgSqlConnection;

        Color colorYavka = System.Drawing.Color.ForestGreen;
        Color colorOtpusk = System.Drawing.Color.Orange;
        Color colorKomandirovka = System.Drawing.Color.Gold;
        Color colorProgul = System.Drawing.Color.Tomato;
        Color colorHollyday = System.Drawing.Color.SteelBlue;
        Color colorSick = System.Drawing.Color.Pink;

        List<List<Int32>> pk_fact = new List<List<Int32>>(); //ключи фактов явки
        List<Pair> modinfied_cells = new List<Pair>(); //координаты измененных ячеек DataGridView

        public Form3()
        {
            InitializeComponent();
            connectPSQL = Connection.get_instance("postgres","Ntcnbhjdfybt_01");
            npgSqlConnection = connectPSQL.get_connect();
            numericUpDown1.Value = DateTime.Now.Year;
            numericUpDown2.Value = DateTime.Now.Month;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form ifrm = Application.OpenForms[0];
            ifrm.Show();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            //List<string> listUnit = new List<string>(); //список подразделений
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

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged; //отписываемся от события
            dataGridView1.Rows.Clear(); //очещаем грид
            pk_fact.Clear(); //очищаем ключи фактов явки 
            blockCellsDays(); //блокируем дни в datagridView
            //ключ подразделения
            NpgsqlCommand com;
            com = new NpgsqlCommand("select \"pk_unit\" from \"Unit\" where \"Unit\".\"Name\" = '" + comboBox1.Text + "'", npgSqlConnection);
            string pk_unit = com.ExecuteScalar().ToString();

            //формируем дату
            string date = numericUpDown1.Value.ToString() + "-" + numericUpDown2.Value.ToString() + "-" + "01";

            //находим ключ табеля
            com = new NpgsqlCommand("SELECT \"pk_time_tracking\" FROM \"TimeTracking\" WHERE \"TimeTracking\".\"pk_unit\" = " + pk_unit + " AND \"TimeTracking\".\"from\" = '" + date + "'", npgSqlConnection);
            string pk_time_tracking;
            try
            {
                pk_time_tracking = com.ExecuteScalar().ToString();
            }
            catch
            {
                MessageBox.Show("Данные не найдены!");
                return;
            }

            //перебор строк найденного табеля
            List<string> pk_personal_card = new List<string>();
            List<string> pk_string_time_tracking = new List<string>();
            com = new NpgsqlCommand("SELECT \"pk_personal_card\",\"pk_string_time_tracking\" FROM \"StringTimeTracking\" WHERE \"StringTimeTracking\".\"pk_time_tracking\" = '" + pk_time_tracking + "'", npgSqlConnection);       
            NpgsqlDataReader reader = com.ExecuteReader();
            if (reader.HasRows)
            {
                foreach (DbDataRecord rec in reader)
                {
                    pk_personal_card.Add(rec.GetInt32(0).ToString()); //ключ личная карточка
                    pk_string_time_tracking.Add(rec.GetInt32(1).ToString()); //ключ строки табеля
                }
            }
            reader.Close();

            for(int i = 0; i < pk_personal_card.Count; i++)
            {
                //получаем ФИО сотрудника
                string fio = "";
                com = new NpgsqlCommand("SELECT \"surname\",\"name\",\"otchestvo\" FROM \"PersonalCard\" WHERE \"PersonalCard\".\"pk_personal_card\" = '" + pk_personal_card[i] + "'", npgSqlConnection);
                reader = com.ExecuteReader();
                if (reader.HasRows)
                {
                    foreach (DbDataRecord rec in reader)
                    {
                        fio += rec.GetString(0) + " ";
                        fio += rec.GetString(1) + " ";
                        fio += rec.GetString(2) + " ";
                    }
                }
                reader.Close();
                //получаем должность сотрудника
                com = new NpgsqlCommand("SELECT \"Position\".\"Name\" FROM \"PeriodPosition\",\"Position\" WHERE \"PeriodPosition\".\"pk_position\" = \"Position\".\"pk_position\" AND \"PeriodPosition\".\"pk_personal_card\" = '" + pk_personal_card[i] + "' AND \"PeriodPosition\".\"DateTo\" is null", npgSqlConnection);
                string name_position = com.ExecuteScalar().ToString();

                //получаем факты явки
                List<Int32> data = new List<Int32>();
                List<string> mark = new List<string>();
                List<Int32> count_of_hours = new List<Int32>();
                com = new NpgsqlCommand("SELECT \"MarkTimeTracking\".\"ShortName\",\"Fact\".\"data\", \"Fact\".\"count_of_hours\",\"Fact\".\"pk_fact\" FROM \"Fact\",\"MarkTimeTracking\" WHERE \"Fact\".\"pk_mark_time_tracking\" = \"MarkTimeTracking\".\"pk_mark_time_tracking\" AND \"Fact\".\"pk_string_time_tracking\" = '" + pk_string_time_tracking[i] + "'", npgSqlConnection);
                reader = com.ExecuteReader();
                if (reader.HasRows)
                {
                    List<Int32> pk = new List<Int32>(); //строка ключей фактов явки
                    for (int k = 0; k < 31; k++) //максимум 31 день
                        pk.Add(0);
                    foreach (DbDataRecord rec in reader)
                    {
                        mark.Add(rec.GetString(0));
                        data.Add(rec.GetDateTime(1).Day);
                        count_of_hours.Add(rec.GetInt32(2));

                        //сохраняем первичный ключ факта явки для будущего редактирования
                        //как бы делаем копию датагрид, где хранятся первичные ключи фактов явки           
                        pk[data[data.Count-1] - 1] = rec.GetInt32(3);
                    }
                    pk_fact.Add(pk); //сохраняем строку ключей фактов явки для будущего редактирования.
                }
                reader.Close();

                //добавляем строку в datagridView
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = fio;
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = name_position;

                for(int j = 0; j < data.Count; j++)
                {
                    // + 1 к индексу, потому что первые два столбца это фио и должность
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[data[j] + 1].Value = mark[j];
                    paintingCells(dataGridView1.Rows.Count - 1, data[j] + 1, mark[j]);
                }

                //добавляем строку часов сотрудника
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = "Количество часов:";
                for (int j = 0; j < data.Count; j++)
                {
                    // + 1 к индексу, потому что первые два столбца это фио и должность
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[data[j] + 1].Value = count_of_hours[j];
                }
            }
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged; //обратно подписываемся на событие 
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            pk_fact.Clear();
        }

        void blockCellsDays() //блокировка дней в datagridView
        {
            Int32 count_day = DateTime.DaysInMonth(Convert.ToInt32(numericUpDown1.Value), Convert.ToInt32(numericUpDown2.Value));

            dataGridView1.Columns[32].Visible = true;
            dataGridView1.Columns[31].Visible = true;
            dataGridView1.Columns[30].Visible = true;
            if (count_day == 30)
            {
                dataGridView1.Columns[32].Visible = false;

            }
            else if (count_day == 28)
            {
                dataGridView1.Columns[32].Visible = false;
                dataGridView1.Columns[31].Visible = false;
                dataGridView1.Columns[30].Visible = false;
            }
            else if (count_day == 29)
            {
                dataGridView1.Columns[32].Visible = false;
                dataGridView1.Columns[31].Visible = false;
            }
        }

        void paintingCells(int row, int colmn, string shifr)
        {
            if (shifr == "Я" || shifr == "Н" || shifr == "РВ" || shifr == "С")
            {
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = colorYavka;
            }
            else if (shifr == "К")
            {
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = colorKomandirovka;
            }
            else if (shifr == "ОТ" || shifr == "ОД")
            {
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = colorOtpusk;
            }
            else if (shifr == "Б")
            {
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = colorSick;
            }
            else if (shifr == "В")
            {
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = colorHollyday;
            }
            else if (shifr == "ПР" || shifr == "НН")
            {
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = colorProgul;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            pk_fact.Clear();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                string shifr = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                if(e.RowIndex % 2 == 0) //если меняется строки с шифрами
                {
                    if (shifr != "Я" && shifr != "Н" && shifr != "РВ" && shifr != "С" && shifr != "К" &&
                        shifr != "ОТ" && shifr != "ОД" && shifr != "Б" && shifr != "ПР" && shifr != "В" && shifr != "НН")
                    {
                        MessageBox.Show("Неверный шифр в ячейке!");
                        dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                        dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                        return;
                    }
                    modinfied_cells.Add(new Pair(e.RowIndex, e.ColumnIndex)); //добавляем координаты изменной ячейки
                    paintingCells(e.RowIndex, e.ColumnIndex, shifr);
                }
                else //если меняется строка с часами
                {
                    for(int i = 0; i < shifr.Length; i++)
                        if( !(shifr[i] >= '0' && shifr[i] <= '9'))
                        {
                            MessageBox.Show("Неверное значение в ячейке часов!");
                            dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                            return;
                        }

                    modinfied_cells.Add(new Pair(e.RowIndex, e.ColumnIndex)); //добавляем координаты изменной ячейки
                }

                
            }
               
        }
    }

    public class Pair
    {
        public Pair(Int32 x, Int32 y) { X = x; Y = y; }
        public Int32 X { get; set; }
        public Int32 Y { get; set; }
    }
}
