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
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace PersonnelDeptApp1
{
    public partial class Form3 : Form
    {
        //Connection connectPSQL;
        NpgsqlConnection npgSqlConnection;

        Color colorYavka = System.Drawing.Color.ForestGreen;
        Color colorOtpusk = System.Drawing.Color.Orange;
        Color colorKomandirovka = System.Drawing.Color.Gold;
        Color colorProgul = System.Drawing.Color.Tomato;
        Color colorHollyday = System.Drawing.Color.SteelBlue;
        Color colorSick = System.Drawing.Color.Pink;

        List<List<Int32>> pk_fact = new List<List<Int32>>(); //ключи фактов явки
        List<Pair> modinfied_cells = new List<Pair>(); //координаты измененных ячеек DataGridView

        string file_name = @"";

        string pk_time_tracking; //ключ табеля

        public Form3()
        {
            InitializeComponent();
			
			//Поменял подключение
            npgSqlConnection = Connection.get_connect();

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
            if (modinfied_cells.Count != 0)
            {
                DialogResult result = MessageBox.Show("Изменения не были сохранены. Вы уверены, что хотите выйти?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.No)
                    return;
            }

			Form ifrm = System.Windows.Forms.Application.OpenForms[System.Windows.Forms.Application.OpenForms.Count - 1];
			ifrm.Show();

		}

        private void Form3_Load(object sender, EventArgs e)
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

            if (Connection.get_role() == Connection.Role.reception)
            {
                menuStrip1.Items[0].Enabled = false;
                dataGridView1.ReadOnly = true;
                button2.Visible = false;
                groupBox3.Visible = false;
                groupBox4.Visible = false;
            }
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
                string name_position ="";
                try
                {
                    name_position = com.ExecuteScalar().ToString();
                }
                catch
                {  
                    com = new NpgsqlCommand("SELECT \"Position\".\"Name\" FROM \"PeriodPosition\",\"Position\" WHERE  \"PeriodPosition\".\"pk_personal_card\" = '" + pk_personal_card[i] + "' AND \"PeriodPosition\".\"DataFrom\" = (select max(\"PeriodPosition\".\"DataFrom\") from \"PeriodPosition\" where \"PeriodPosition\".\"pk_personal_card\" = '" + pk_personal_card[i] + "') AND \"PeriodPosition\".\"pk_position\" = \"Position\".\"pk_position\"", npgSqlConnection);
                    reader = com.ExecuteReader();
                    if (reader.HasRows)
                    {
                        foreach (DbDataRecord rec in reader)
                        {
                            name_position = rec.GetString(0);
                        }
                    }
                    reader.Close();
                }
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
            button15.Enabled = true;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

            if (modinfied_cells.Count != 0)
            {
                DialogResult result = MessageBox.Show("Изменения не были сохранены. Вы уверены, что хотите посмотреть другой табель?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.No)
                {
                    numericUpDown2.ValueChanged -= numericUpDown2_ValueChanged;
                    numericUpDown2.Value++;
                    numericUpDown2.ValueChanged += numericUpDown2_ValueChanged;
                    return;
                }
            }
            modinfied_cells.Clear();
            button2.Enabled = false;

            dataGridView1.Rows.Clear();
            pk_fact.Clear();
            button15.Enabled = false;
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
            if (modinfied_cells.Count != 0)
            {
                DialogResult result = MessageBox.Show("Изменения не были сохранены. Вы уверены, что хотите посмотреть другой табель?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.No)
                {
                    numericUpDown1.ValueChanged -= numericUpDown1_ValueChanged;
                    numericUpDown1.Value++;
                    numericUpDown1.ValueChanged += numericUpDown1_ValueChanged;
                    return;
                }
            }
            modinfied_cells.Clear();
            button2.Enabled = false;

            dataGridView1.Rows.Clear();
            button15.Enabled = false;
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
            button15.Enabled = false;
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

        private void button15_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = "Без названия";
            sfd.Filter = "Excel Files(.xls)|*.xls| Excel Files(.xlsx) | *.xlsx | Excel Files(*.xlsm) | *.xlsm";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                file_name = @sfd.FileName;
                if (file_name == "")
                {
                    MessageBox.Show("Имя файла не может быть пустым");
                    return;
                }
                try
                {
                    File.Delete(file_name);
                }
                catch
                {
                    MessageBox.Show("Невозможно перезаписать файл. Файл используется в другом процессе.");
                    return;
                }
                File.Copy(Directory.GetCurrentDirectory() + "\\tabel.xls", file_name);
                writeToExcel();

                MessageBox.Show("Табель успешно экспортирован! Дождитесь открытия файла.");
                Process.Start(file_name);
            }

 
        }

        public void writeToExcel()
        {
            // Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(file_name, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            //шапка
            ObjWorkSheet.Cells[7, "A"] = "Городская больница №5"; //наименование организации
            ObjWorkSheet.Cells[9, "A"] = comboBox1.Text; //подразделение
            ObjWorkSheet.Cells[7, "GV"] = "00034237";

            //ключ подразделения
            NpgsqlCommand com;
            com = new NpgsqlCommand("select * from \"TimeTracking\" where \"TimeTracking\".\"pk_time_tracking\" = '" + pk_time_tracking + "'", npgSqlConnection);
            NpgsqlDataReader reader = com.ExecuteReader();
            if (reader.HasRows)
            {
                foreach (DbDataRecord rec in reader)
                {
                    ObjWorkSheet.Cells[13, "EE"] = rec.GetString(1);
                    ObjWorkSheet.Cells[13, "EU"] = rec.GetDateTime(2).ToString("dd.MM.yyyy");
                    ObjWorkSheet.Cells[13, "FN"] = rec.GetDateTime(4).ToString("dd.MM.yyyy");
                    ObjWorkSheet.Cells[13, "FY"] = rec.GetDateTime(3).ToString("dd.MM.yyyy");
                }
            }
            reader.Close();

            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[2];
            //строки табеля
            int first_half_hours, second_half_hours, first_half_days, second_half_days;
            int overtime_hours, night_hours, holliday_hours, hollidays_days;
            int neyavka_days, prichina_days;
            for (int i = 0, j = 9; i  < dataGridView1.RowCount; i++, j++)
            {
                if ( i % 2 == 0) //строка с шифром
                {
                    first_half_days = 0; second_half_days = 0;
                    hollidays_days = 0; neyavka_days = 0; prichina_days = 0;
                    ObjWorkSheet.Cells[j, "A"] = i / 2  + 1; //номер по порядку
                    ObjWorkSheet.Cells[j, "F"] = dataGridView1.Rows[i].Cells[0].Value.ToString(); //имя 
                    com = new NpgsqlCommand("select \"pk_personal_card\" from \"StringTimeTracking\" where \"StringTimeTracking\".\"pk_string_time_tracking\" = '" + dataGridView1.Rows[i].Cells[33].Value.ToString() + "'", npgSqlConnection);
                    ObjWorkSheet.Cells[j, "V"] = com.ExecuteScalar().ToString(); //табельный номер

                    List<string> word_indx = new List<string>() {"AG", "AK", "AO", "AS", "AW", "BA", "BE", "BI", "BM", "BQ", "BU", "BY", "CC", "CG", "CK"};
                    for (int k = 0; k < 15; k++) //первая половина месяца
                    {
                        if (dataGridView1.Rows[i].Cells[k+2].Value != null)
                            if (dataGridView1.Rows[i].Cells[k+2].Value.ToString() != "")
                            {
                                string shfr = dataGridView1.Rows[i].Cells[k + 2].Value.ToString();
                                ObjWorkSheet.Cells[j, word_indx[0]] = shfr;
                                
                                if (shfr == "Я" || shfr == "Н" || shfr == "РВ" || shfr == "С")
                                    first_half_days++;
                                else if (shfr == "В")
                                    hollidays_days++;
                                else if (shfr == "Б" || shfr == "НН" || shfr == "ПР")
                                {
                                    neyavka_days++;
                                    if (shfr == "Б")
                                        prichina_days++;
                                }
                            }
                        word_indx.RemoveAt(0);
                    }
                    ObjWorkSheet.Cells[j, "CO"] = first_half_days;

                    word_indx.AddRange( new string[] { "CV", "CZ", "DD", "DH", "DL", "DP", "DT", "DX", "EB", "EF", "EJ", "EN", "ER", "EV", "EZ", "FD" });

                    for (int k = 15; k < 31; k++) //первая половина месяца
                    {
                        if (dataGridView1.Rows[i].Cells[k + 2].Value != null && dataGridView1.Columns[k+2].Visible == true)
                            if (dataGridView1.Rows[i].Cells[k + 2].Value.ToString() != "")
                            {
                                string shfr = dataGridView1.Rows[i].Cells[k + 2].Value.ToString();
                                ObjWorkSheet.Cells[j, word_indx[0]] = shfr;
                                if (shfr == "Я" || shfr == "Н" || shfr == "РВ" || shfr == "С")
                                    second_half_days++;
                                else if (shfr == "В")
                                    hollidays_days++;
                                else if (shfr == "Б" || shfr == "НН" || shfr == "ПР")
                                {
                                    neyavka_days++;
                                    if (shfr == "Б")
                                        prichina_days++;
                                }
                            }
                        word_indx.RemoveAt(0);
                    }
                    ObjWorkSheet.Cells[j, "FH"] = second_half_days;

                    ObjWorkSheet.Cells[j, "FO"] = (first_half_days + second_half_days).ToString();
                    ObjWorkSheet.Cells[j, "HE"] = neyavka_days.ToString();
                    ObjWorkSheet.Cells[j, "HM"] = "Б";
                    ObjWorkSheet.Cells[j, "HS"] = prichina_days.ToString();
                    ObjWorkSheet.Cells[j, "IA"] = hollidays_days;
                }
                else //строка с часами
                {
                    first_half_hours = 0; second_half_hours = 0;
                    overtime_hours = 0; night_hours = 0; holliday_hours = 0;
                    
                    List<string> word_indx = new List<string>() { "AG", "AK", "AO", "AS", "AW", "BA", "BE", "BI", "BM", "BQ", "BU", "BY", "CC", "CG", "CK" };
                    for (int k = 0; k < 15; k++) //первая половина месяца
                    {
                        if (dataGridView1.Rows[i].Cells[k + 2].Value != null)
                            if (dataGridView1.Rows[i].Cells[k + 2].Value.ToString() != "")
                            {
                                int cnt_hours = Convert.ToInt32(dataGridView1.Rows[i].Cells[k + 2].Value);
                                ObjWorkSheet.Cells[j, word_indx[0]] = cnt_hours.ToString();
                                first_half_hours += cnt_hours;

                                if (dataGridView1.Rows[i - 1].Cells[k + 2].Value.ToString() == "С")
                                    overtime_hours += cnt_hours;
                                else if (dataGridView1.Rows[i-1].Cells[k + 2].Value.ToString() == "Н")
                                    night_hours += cnt_hours;
                                else if (dataGridView1.Rows[i - 1].Cells[k + 2].Value.ToString() == "РВ")
                                    holliday_hours += cnt_hours;
                            }
                        word_indx.RemoveAt(0);
                    }
                    ObjWorkSheet.Cells[j, "CO"] = first_half_hours;

                    word_indx.AddRange(new string[] { "CV", "CZ", "DD", "DH", "DL", "DP", "DT", "DX", "EB", "EF", "EJ", "EN", "ER", "EV", "EZ", "FD" });

                    for (int k = 15; k < 31; k++) //первая половина месяца
                    {
                        if (dataGridView1.Rows[i].Cells[k + 2].Value != null && dataGridView1.Columns[k + 2].Visible == true)
                            if (dataGridView1.Rows[i].Cells[k + 2].Value.ToString() != "")
                            {
                                int cnt_hours = Convert.ToInt32(dataGridView1.Rows[i].Cells[k + 2].Value);
                                ObjWorkSheet.Cells[j, word_indx[0]] = cnt_hours.ToString();
                                second_half_hours += cnt_hours;

                                if (dataGridView1.Rows[i - 1].Cells[k + 2].Value.ToString() == "С")
                                    overtime_hours += cnt_hours;
                                else if (dataGridView1.Rows[i - 1].Cells[k + 2].Value.ToString() == "Н")
                                    night_hours += cnt_hours;
                                else if (dataGridView1.Rows[i - 1].Cells[k + 2].Value.ToString() == "РВ")
                                    holliday_hours += cnt_hours;
                            }
                        word_indx.RemoveAt(0);
                    }
                    ObjWorkSheet.Cells[j, "FH"] = second_half_hours;

                    ObjWorkSheet.Cells[j - 1, "FV"] = (first_half_hours + second_half_hours).ToString();
                    ObjWorkSheet.Cells[j - 1, "GC"] = overtime_hours.ToString();
                    ObjWorkSheet.Cells[j - 1, "GJ"] = night_hours.ToString();
                    ObjWorkSheet.Cells[j - 1, "GQ"] = holliday_hours.ToString();

                    ObjWorkSheet.Cells[j, "HE"] = (Convert.ToInt32(ObjWorkSheet.Cells[j - 1, "HE"].Text) * 8).ToString();
                    ObjWorkSheet.Cells[j, "HM"] = "Б";
                    ObjWorkSheet.Cells[j, "HS"] = (Convert.ToInt32(ObjWorkSheet.Cells[j - 1, "HS"].Text) * 8).ToString();

                }

            }

            //закрытие документа
            ObjWorkBook.Close(true);
            ObjExcel.Quit();
            ObjExcel = null;
            ObjWorkBook = null;
            ObjWorkSheet = null;
            GC.Collect();
        }

        private void управлениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form204 FormAddTabel = new Form204();

            DialogResult result = FormAddTabel.ShowDialog(this);

            
            if (result == DialogResult.OK)
            {
                //открытие созданного табеля
                comboBox1.Text = FormAddTabel.comboBox1.Text;
                numericUpDown1.Value = FormAddTabel.numericUpDown1.Value;
                numericUpDown2.Value = FormAddTabel.numericUpDown2.Value;
                button1_Click(null, null);
            }
            
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
