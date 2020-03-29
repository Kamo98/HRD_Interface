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
            //button2.Enabled = false;
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
            if (modinfied_cells.Count != 0)
            {
                DialogResult result = MessageBox.Show("Изменения не были сохранены. Вы уверены, что хотите выйти?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                    return;
            }
            
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
                List<Int32> pk = new List<Int32>(); //строка ключей фактов явки
                for (int k = 0; k < dataGridView1.ColumnCount; k++)
                    pk.Add(-1);
                com = new NpgsqlCommand("SELECT \"MarkTimeTracking\".\"ShortName\",\"Fact\".\"data\", \"Fact\".\"count_of_hours\",\"Fact\".\"pk_fact\" FROM \"Fact\",\"MarkTimeTracking\" WHERE \"Fact\".\"pk_mark_time_tracking\" = \"MarkTimeTracking\".\"pk_mark_time_tracking\" AND \"Fact\".\"pk_string_time_tracking\" = '" + pk_string_time_tracking[i] + "'", npgSqlConnection);
                reader = com.ExecuteReader();
                if (reader.HasRows)
                {           
                    foreach (DbDataRecord rec in reader)
                    {
                        mark.Add(rec.GetString(0));
                        data.Add(rec.GetDateTime(1).Day);
                        count_of_hours.Add(rec.GetInt32(2));

                        //сохраняем первичный ключ факта явки для будущего редактирования
                        //как бы делаем копию датагрид, где хранятся первичные ключи фактов явки           
                        pk[data[data.Count-1] + 1] = rec.GetInt32(3);
                    }
                }
                pk_fact.Add(pk); //сохраняем строку ключей фактов явки для будущего редактирования.               
                pk_fact.Add(new List<Int32>());
                reader.Close();

                //добавляем строку в datagridView
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = fio;
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = name_position;
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[dataGridView1.ColumnCount - 1].Value = pk_string_time_tracking[i];

                for (int j = 0; j < data.Count; j++)
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
            button2.Enabled = false;
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
            else
                dataGridView1.Rows[row].Cells[colmn].Style.BackColor = System.Drawing.Color.White;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            button2.Enabled = true;
            if (e.RowIndex != -1)
            {
                string shifr;

                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
                    shifr = "";
                else shifr = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                if (e.RowIndex % 2 == 0) //если меняется строки с шифрами
                {
                    dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                    if (shifr != "Я" && shifr != "Н" && shifr != "РВ" && shifr != "С" && shifr != "К" &&
                        shifr != "ОТ" && shifr != "ОД" && shifr != "Б" && shifr != "ПР" && shifr != "В" && shifr != "НН" && shifr != "")
                    {
                        MessageBox.Show("Неверный шифр в ячейке!");                   
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                        dataGridView1.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value = "0";
                        dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                        return;
                        
                    }
                    if (shifr != "Я" && shifr != "Н" && shifr != "РВ" && shifr != "С" && shifr != "")
                    {
                        dataGridView1.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value = 0;
                    }
                    else if(shifr == "")
                    {
                        dataGridView1.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value = "";
                    }

                    if (dataGridView1.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value == null)
                    {
                        dataGridView1.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value = 0;
                    }
                    dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;

                    modinfied_cells.Add(new Pair(e.RowIndex, e.ColumnIndex)); //добавляем координаты изменной ячейки
                    paintingCells(e.RowIndex, e.ColumnIndex, shifr);
                }
                else //если меняется строка с часами
                {
                    dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                    for (int i = 0; i < shifr.Length; i++)
                        if( !(shifr[i] >= '0' && shifr[i] <= '9'))
                        {
                            MessageBox.Show("Неверное значение в ячейке часов!");    
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                            return;
                        }

                    if (dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value == null)
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                        dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                        return;
                    }
                    else if (dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() != "Я"
                        && dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() != "Н"
                         && dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() != "РВ"
                          && dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() != "С")
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                    }
                    else if (shifr == "")
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                    }
                    dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;

                    modinfied_cells.Add(new Pair(e.RowIndex - 1, e.ColumnIndex)); //добавляем координаты изменной ячейки
                }         
            }
               
        }

        private void button2_Click(object sender, EventArgs e)
        {
            NpgsqlCommand com;

            button2.Enabled = false;

            //отсеиваем лишние повторяющиеся изменения
            modinfied_cells = modinfied_cells.Distinct(new PairComparer()).ToList();

            for (int i = 0; i < modinfied_cells.Count; i++)
            {
                Int32 x = modinfied_cells[i].X, y = modinfied_cells[i].Y;
                if (pk_fact[x][y] == -1) //если ключ не существует, то insert
                {
                    //определяем ключ метки шифра
                    com = new NpgsqlCommand("SELECT \"pk_mark_time_tracking\" FROM \"MarkTimeTracking\" WHERE \"MarkTimeTracking\".\"ShortName\" = '" + dataGridView1.Rows[x].Cells[y].Value + "'", npgSqlConnection);
                    string pk_mark = com.ExecuteScalar().ToString();
                    //ключ строки в табеле
                    string pk_string = dataGridView1.Rows[x].Cells[dataGridView1.ColumnCount - 1].Value.ToString();
                    //дата
                    string data = numericUpDown1.Value.ToString() + "-" + numericUpDown2.Value.ToString() + "-" + (y - 1).ToString();
                    //часы
                    string hours = dataGridView1.Rows[x + 1].Cells[y].Value.ToString();
                    if (hours == "")
                        hours = "0";

                    com = new NpgsqlCommand("INSERT INTO \"Fact\"(\"pk_string_time_tracking\", \"pk_mark_time_tracking\", \"data\", \"count_of_hours\") VALUES('" + pk_string + "', '" + pk_mark + "', '" + data + "', '" + hours + "')", npgSqlConnection);
                    com.ExecuteNonQuery();
                }
                else //если ключ существует, то update или delete
                {
                    if(dataGridView1.Rows[x].Cells[y].Value == null) //удаление
                    {
                        com = new NpgsqlCommand("DELETE FROM \"Fact\"  WHERE \"Fact\".\"pk_fact\" = '" + pk_fact[x][y] + "'", npgSqlConnection);
                        com.ExecuteNonQuery();
                    }
                    else if (dataGridView1.Rows[x].Cells[y].Value.ToString() == "")
                    {
                        com = new NpgsqlCommand("DELETE FROM \"Fact\"  WHERE \"Fact\".\"pk_fact\" = '" + pk_fact[x][y] + "'", npgSqlConnection);
                        com.ExecuteNonQuery();
                    }
                    else
                    {
                        //определяем ключ метки шифра
                        com = new NpgsqlCommand("SELECT \"pk_mark_time_tracking\" FROM \"MarkTimeTracking\" WHERE \"MarkTimeTracking\".\"ShortName\" = '" + dataGridView1.Rows[x].Cells[y].Value + "'", npgSqlConnection);
                        string pk_mark = com.ExecuteScalar().ToString();
                        //часы
                        string hours = dataGridView1.Rows[x + 1].Cells[y].Value.ToString();
                        if (hours == "")
                            hours = "0";
                        com = new NpgsqlCommand("UPDATE \"Fact\" SET \"pk_mark_time_tracking\" = '" + pk_mark + "', \"count_of_hours\" = '" + hours + "' WHERE \"Fact\".\"pk_fact\" = '" + pk_fact[x][y] + "'", npgSqlConnection);
                        com.ExecuteNonQuery();
                    }   
                }
            }
            MessageBox.Show("Изменения успешно сохранены!");
            //обновляем, чтобы подгрузились ключи в датагрид
            button1_Click(null, null);
            modinfied_cells.Clear();         
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CellFilling("Я");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CellFilling("Б");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CellFilling("В");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            CellFilling("НН");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            CellFilling("ПР");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CellFilling("Н");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            CellFilling("РВ");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            CellFilling("С");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            CellFilling("К");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            CellFilling("ОТ");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            CellFilling("ОД");
        }

        void CellFilling(string shifr)
        {
            if (dataGridView1.SelectedCells.Count == 0)
            {
                MessageBox.Show("Сперва выделите ячейки!");
            }
            for (int i = 0; i < dataGridView1.SelectedCells.Count; i++)
                if (dataGridView1.CurrentCellAddress.X != 0 && dataGridView1.CurrentCellAddress.X != 1)
                    dataGridView1.SelectedCells[i].Value = shifr;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            CellFilling(numericUpDown3.Value.ToString());
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
    }

    public class Pair
    {
        public Pair(Int32 x, Int32 y) { X = x; Y = y; }
        public Int32 X { get; set; }
        public Int32 Y { get; set; }
    }

    class PairComparer : IEqualityComparer<Pair>
    {
        public bool Equals(Pair first, Pair second)
        {
            return first.X == second.X && first.Y == second.Y;
        }

        public int GetHashCode(Pair x)
        {
            return (x.X + "_" + x.Y).GetHashCode();
        }
    }
}
