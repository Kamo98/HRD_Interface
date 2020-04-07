using HRD_GenerateData;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace PersonnelDeptApp1
{
    public partial class Form4 : Form
    {
        /*
         * Описание:
         *      >>  При загрузке формы подгружаются список отделов и список должностей соответствующих выбранному отделу
         *      >>  При смене отдела список должностей обновляется
         *      >>  При смене вкладки (таба) все формы сбрасываются в дефолтное состояние
         *      >>  Для выбора сотрудника начните вводить ФИО, 
         *          затем выберите сотрудника из появившегося списка, используя стрелки клавиатуры (↓ и ↑),
         *          нажмите Enter для подтверждения выбора.
         *          В поле ФИО будет внесено ФИО и номер карты (для точной идентификации сотрудника).
         *      >>  Иногда вылетает ошибка чтения из потока (скорее всего пропадает соединение с БД)
         *      >>  На данный момент в БД ничего не вносится и ничего связанного с приказами (например, номер договора) не достается
         */
        enum OrderType{
            Hire,
            Fire,
            Move
        }
        string dateFormat = "dd-MM-yyyy";
        Dictionary<OrderType, string> orderDict = new Dictionary<OrderType, string>();
        string orderNum;
        System.Windows.Forms.ListBox employeesVars = new System.Windows.Forms.ListBox();
        BindingList<Employee> employees = new BindingList<Employee>();
        BindingList<Department> departments = new BindingList<Department>();
        BindingList<Occupation> occupations = new BindingList<Occupation>();
        Connection connection = Connection.get_instance("postgres", "Ntcnbhjdfybt_01");
        public Form4()
        {
            InitializeComponent();
            FillDepts();
            
            orderNum = GenereteOrderNum(7);
            hireDocNum.Text = orderNum;
            employeesVars.DataSource = employees;
            employeesVars.DisplayMember = "FIO";
            employeesVars.ValueMember = "Id";

            orderDict.Add(OrderType.Fire, "Увольнение");
            orderDict.Add(OrderType.Hire, "Приём");
            orderDict.Add(OrderType.Move, "Перевод");
        }

        private void FillDepts()
        {
            try
            {
                string sql = "select * from \"Unit\";";
                if (connection.get_connect() == null)
                    throw new NullReferenceException("Не удалось подключиться к базе данных");
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                foreach (DbDataRecord record in reader)
                {
                    Department newDept = new Department((int)record["pk_unit"], (string)record["Name"]);
                    departments.Add(newDept);

                }
                reader.Close();

                hireDepartment.DataSource = moveDepartmentNew.DataSource = departments;
                hireDepartment.DisplayMember = moveDepartmentNew.DisplayMember = "Name";
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex) {
                MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FillOccups(object sender, EventArgs e)
        {
            occupations.Clear();

            try
            {
                string sql = "select \"pos\".\"pk_position\", \"pos\".\"Name\", \"pos\".\"Rate\""
                                + " from \"Position\" as \"pos\", \"Unit\" as \"un\""
                                + " where \"pos\".\"pk_unit\" = \"un\".\"pk_unit\" and \"un\".\"pk_unit\" = " + ((sender as ComboBox).SelectedItem as Department).Id + ";";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                foreach (DbDataRecord record in reader)
                {
                    occupations.Add(new Occupation((int)record["pk_position"], (string)record["Name"], (decimal)record["Rate"]));
                }
                reader.Close();

                hireOccup.DataSource = moveOccupNew.DataSource = occupations;
                hireOccup.DisplayMember = moveOccupNew.DisplayMember = "Name";

                if (orderTab.SelectedTab == hirePage && hireOccup.SelectedItem != null)
                    hireTarif.Text = moveTarif.Text = ((Occupation)(hireOccup.SelectedItem)).Tarif.ToString();
                else if (orderTab.SelectedTab == movePage && moveOccupNew.SelectedItem != null)
                    hireTarif.Text = moveTarif.Text = ((Occupation)(moveOccupNew.SelectedItem)).Tarif.ToString();
            }
            catch (Exception ex) {
                MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        /*
         * Недостатки текущего автокомплита:
         *      ----    Долгая обработка (будет особенно ощутимо на больших массивах данных).
         *              Связано это с тем, что список подходящих данных постоянно заново подгружается из БД
         */
        private void FIO_Autocomplete(object sender, EventArgs e)
        {
            if (connection == null)
                return;

            employees.Clear();

            if ((sender as RichTextBox).Text.Equals(""))
                return;
            string sql = "select * from getlist_by_substring('" + (sender as RichTextBox).Text.Split('#')[0].Trim() + "');";
            NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
            NpgsqlDataReader reader = command.ExecuteReader();
            foreach (DbDataRecord record in reader)
            {
                Employee newEmp = new Employee((int)record[1], record[0].ToString());
                employees.Add(newEmp);
            }

            reader.Close();
        }

        private void KeyPressOnFIOField(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                employeesVars.SelectedIndex = Math.Min(employeesVars.Items.Count - 1, employeesVars.SelectedIndex + 1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (employeesVars.Items.Count != 0)
                    employeesVars.SelectedIndex = Math.Max(0, employeesVars.SelectedIndex - 1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                Employee selected = (Employee)employeesVars.SelectedItem;
                if (selected != null)
                    (sender as RichTextBox).Text = selected.FIO + "  #" + selected.Id;
                (sender as RichTextBox).Parent.Focus();
                if ((sender as RichTextBox) == fireFIO || (sender as RichTextBox) == moveFIO) {
                    string sql =
                        "select \"dep\".\"Name\", \"pos\".\"Name\", \"pos\".\"Rate\", \"doc\".\"Number_work_doc\", \"doc\".\"Work_doc_date\"" +
                        " from \"String_order\" as \"doc\", \"Position\" as \"pos\", \"Unit\" as \"dep\", \"PeriodPosition\" as \"pp\"" +
                        " where   \"pp\".\"DateTo\" is null and" +
                                " \"pp\".\"pk_personal_card\" = " + selected.Id +
                                " and \"pp\".\"pk_position\" = \"pos\".\"pk_position\" and" +
                                " \"pp\".\"pk_move_order\" = \"doc\".\"pk_string_order\" and" +
                                " \"pos\".\"pk_unit\" = \"dep\".\"pk_unit\"";
                    NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                    NpgsqlDataReader reader = command.ExecuteReader();
                    if (!reader.HasRows)
                    {
                        reader.Close();
                        MessageBox.Show("Не существует действующего договора с данным сотрудником!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Handled = true;
                        return;
                    }
                    else {
                        if ((sender as RichTextBox) == fireFIO)
                        {
                            reader.Read();
                            fireDepartment.Text = reader.GetString(0);
                            fireOccup.Text = reader.GetString(1);
                            fireTarif.Text = reader.GetDecimal(2).ToString();
                            fireContractNum.Text = reader.GetString(3);
                            fireContractDate.Value = reader.GetDateTime(4);
                        }
                        else {
                            reader.Read();
                            moveDepartmentOld.Text = reader.GetValue(0).ToString();
                            moveOccupOld.Text = reader.GetValue(1).ToString();
                        }
                    }
                    reader.Close();
                }
                e.Handled = true;
            }
        }

        private void InputLimitOnlyLetters(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar))
                e.Handled = true;
        }

        private void FIO_fieldInFocus(object sender, EventArgs e)
        {
            employeesVars.Location = new System.Drawing.Point(
               (sender as RichTextBox).Location.X,
               (sender as RichTextBox).Location.Y + (sender as RichTextBox).Height
               );
            employeesVars.Size = new Size(200, 50);
            employeesVars.Parent = (sender as RichTextBox).Parent;
            employeesVars.BringToFront();
            employeesVars.Show();
        }
        private void FIO_fieldDropFocus(object sender, EventArgs e)
        {
            employeesVars.Hide();
        }

        private void ChangeOccupationsItem(object sender, EventArgs e)
        {
            if (orderTab.SelectedTab == hirePage && hireOccup.SelectedItem != null)
                hireTarif.Text = moveTarif.Text = ((Occupation)(hireOccup.SelectedItem)).Tarif.ToString();
            else if (orderTab.SelectedTab == movePage && moveOccupNew.SelectedItem != null)
                hireTarif.Text = moveTarif.Text = ((Occupation)(moveOccupNew.SelectedItem)).Tarif.ToString();
        }

        private void addIntoHireOrderBTN_Click(object sender, EventArgs e)
        {
            try
            {
                string sql = "Select * from \"PeriodPosition\"" 
                    + " where \"pk_personal_card\" = " + (employeesVars.SelectedItem as Employee).Id + " and \"DateTo\" is null";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows) {
                    reader.Close();
                    MessageBox.Show("Существует действующий договор с данным сотрудником!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                reader.Close();
                hireTable.Rows.Add(
                    (employeesVars.SelectedItem as Employee).FIO,
                    (employeesVars.SelectedItem as Employee).Id,
                    (hireDepartment.SelectedItem as Department).Name,
                    (hireOccup.SelectedItem as Occupation).Name,
                    (hireOccup.SelectedItem as Occupation).Tarif,
                    hireContractNum.Text,
                    hireContractDate.Value.ToString(dateFormat),
                    startWork.Value.ToString(dateFormat)
                    );

                hireFIO.Text = "";
                employees.Clear();
                hireContractNum.Text = "";
                hireContractDate.Value = startWork.Value = DateTime.Now;
                hireDepartment.SelectedIndex = 0;
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show("Введены не все данные. Введите все необходимые данные!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void orderTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            orderNum = GenereteOrderNum(7);
            HireTabReset();
            MoveTabReset();
            FireTabReset();
        }

        private void HireTabReset() {
            hireFIO.Text = "";
            employees.Clear();
            hireContractNum.Text = "";
            hireDocNum.Text = orderNum;
            hireDocDate.Value = hireContractDate.Value = startWork.Value = DateTime.Now;
            if (hireDepartment.SelectedIndex != 0)
                hireDepartment.SelectedIndex = 0;
            hireTable.Rows.Clear();
        }
        private void MoveTabReset()
        {
            moveFIO.Text = "";
            employees.Clear();
            moveContractNum.Text = "";
            moveDocNum.Text = orderNum;
            moveDocDate.Value = moveContractDate.Value = DateTime.Now;
            moveDepartmentOld.Text = "";
            moveOccupOld.Text = "";
            fireReason.Text = "";
            moveTable.Rows.Clear();
        }
        private void FireTabReset()
        {
            fireFIO.Text = "";
            employees.Clear();
            fireContractNum.Text = "";
            fireDocNum.Text = orderNum;
            fireDocDate.Value = moveContractDate.Value = DateTime.Now;
            fireDepartment.Text = "";
            fireOccup.Text = "";
            fireTarif.Text = "";
            fireTable.Rows.Clear();
        }

        private void addIntoFireOrderBTN_Click(object sender, EventArgs e)
        {
            try
            {
                string sql = "Select * from \"PeriodPosition\""
                    + " where \"pk_personal_card\" = " + (employeesVars.SelectedItem as Employee).Id + " and \"DateTo\" is null";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                if (!reader.HasRows)
                {
                    reader.Close();
                    MessageBox.Show("Не существует действующего договора с данным сотрудником!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                reader.Close();
                fireTable.Rows.Add(
                    (employeesVars.SelectedItem as Employee).FIO,
                    (employeesVars.SelectedItem as Employee).Id,
                    fireDepartment.Text,
                    fireOccup.Text,
                    fireContractNum.Text,
                    fireContractDate.Value.ToString(dateFormat),
                    fireReason.Text,
                    endWork.Value.ToString(dateFormat)
                    );

                fireFIO.Text = "";
                employees.Clear();
                fireContractNum.Text = "";
                fireContractDate.Value = endWork.Value = DateTime.Now;
                fireDepartment.Text = fireReason.Text = fireOccup.Text = fireTarif.Text = "";
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show("Введены не все данные. Введите все необходимые данные!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void addIntoMoveOrderBTN_Click(object sender, EventArgs e)
        {
            try
            {
                string sql = "Select * from \"PeriodPosition\""
                   + " where \"pk_personal_card\" = " + (employeesVars.SelectedItem as Employee).Id + " and \"DateTo\" is null";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                if (!reader.HasRows)
                {
                    reader.Close();
                    MessageBox.Show("Не существует действующего договора с данным сотрудником!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                reader.Close();
                moveTable.Rows.Add(
                    (employeesVars.SelectedItem as Employee).FIO,
                    (employeesVars.SelectedItem as Employee).Id,
                    moveDepartmentOld.Text,
                    (moveDepartmentNew.SelectedItem as Department).Name,
                    moveOccupOld.Text,
                    (moveOccupNew.SelectedItem as Occupation).Name,
                    moveTarif.Text,
                    moveContractNum.Text,
                    moveContractDate.Value.ToString(dateFormat),
                    newPositionDate.Value.ToString(dateFormat)
                    );

                moveFIO.Text = "";
                employees.Clear();
                moveContractNum.Text = "";
                moveContractDate.Value = DateTime.Now;
                moveDepartmentOld.Text = moveOccupOld.Text = "";
                if (moveDepartmentNew.SelectedIndex != 0)
                    moveDepartmentNew.SelectedIndex = 0;
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show("Введены не все данные. Введите все необходимые данные!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void addMoveOrder_Click(object sender, EventArgs e)
        {
            List<int> orderStrings = new List<int>();
            int order = 0;
            try
            {
                if (moveTable.Rows.Count == 0)
                    throw new EmptyTableError("Таблица приказа пуста!");
                order = CreateOrder(OrderType.Move, moveDocNum.Text, moveDocDate.Value.ToString(dateFormat));
                if (order == -1)
                    return;
                for (int i = 0; i < moveTable.Rows.Count; i++) {
                    int oneString = CreateOrderString(
                                            order,
                                            moveTable.Rows[i].Cells[7].Value.ToString(),
                                            moveTable.Rows[i].Cells[8].Value.ToString(),
                                            "");
                    if (oneString == -1)
                        throw new DbInsertErrorException();
                    orderStrings.Add(oneString);
                    ClosePeriodPosition(
                        (int)moveTable.Rows[i].Cells[1].Value, 
                        moveTable.Rows[i].Cells[9].Value.ToString());
                    CreatePeriodPosition(
                        oneString, 
                        (int)moveTable.Rows[i].Cells[1].Value, 
                        GetPositionPKByName(moveTable.Rows[i].Cells[5].Value.ToString()), 
                        moveTable.Rows[i].Cells[9].Value.ToString());
                }

                if (MessageBox.Show("Приказ успешно добавлен!\nСохранить в Excel-файл?", "Успех", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    MoveToExcel();
                MoveTabReset();
            }
            catch (DbInsertErrorException ex) {
                string sql;
                foreach (int one in orderStrings) {
                    sql = "Delete from \"String_order\" where \"pk_string_order\" = " + one;
                    new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                }
                sql = "Delete from \"Order\" where \"pk_order\" = " + order;
                new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                return;
            }
            catch (EmptyTableError ETerr)
            {
                MessageBox.Show(ETerr.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex) {
                MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void addFireOrder_Click(object sender, EventArgs e)
        {
            List<int> orderStrings = new List<int>();
            int order = 0;
            try
            {
                if (fireTable.Rows.Count == 0)
                    throw new EmptyTableError("Таблица приказа пуста!");
                order = CreateOrder(OrderType.Fire, fireDocNum.Text, fireDocDate.Value.ToString(dateFormat));
                if (order == -1)
                    return;
                for (int i = 0; i < fireTable.Rows.Count; i++)
                {
                    int oneString = CreateOrderString(
                                            order,
                                            fireTable.Rows[i].Cells[4].Value.ToString(),
                                            fireTable.Rows[i].Cells[5].Value.ToString(),
                                            fireTable.Rows[i].Cells[6].Value.ToString());
                    if (oneString == -1)
                        throw new DbInsertErrorException();
                    orderStrings.Add(oneString);
                    string sql = "Update \"PeriodPosition\" set \"pk_fire_order_string\" = " + oneString
                       + " where \"pk_personal_card\" = " + (int)fireTable.Rows[i].Cells[1].Value + " and \"DateTo\" is null;";
                    new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                    ClosePeriodPosition(
                        (int)fireTable.Rows[i].Cells[1].Value,
                        fireTable.Rows[i].Cells[7].Value.ToString());
                }

                if (MessageBox.Show("Приказ успешно добавлен!\nСохранить в Excel-файл?", "Успех", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    FireToExcel();
                FireTabReset();
                
            }
            catch (DbInsertErrorException ex)
            {
                string sql;
                foreach (int one in orderStrings)
                {
                    sql = "Delete from \"String_order\" where \"pk_string_order\" = " + one;
                    new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                }
                sql = "Delete from \"Order\" where \"pk_order\" = " + order;
                new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                return;
            }
            catch (EmptyTableError ETerr) {
                MessageBox.Show(ETerr.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void addHireOrderBTN_Click(object sender, EventArgs e)
        {
            List<int> orderStrings = new List<int>();
            int order = 0;
            try
            {
                if (hireTable.Rows.Count == 0)
                    throw new EmptyTableError("Таблица приказа пуста!");
                
                order = CreateOrder(OrderType.Hire, hireDocNum.Text, hireDocDate.Value.ToString(dateFormat));
                if (order == -1)
                    return;
                for (int i = 0; i < hireTable.Rows.Count; i++)
                {
                    int oneString = CreateOrderString(
                                            order,
                                            hireTable.Rows[i].Cells[5].Value.ToString(),
                                            hireTable.Rows[i].Cells[6].Value.ToString(),
                                            "");
                    if (oneString == -1)
                        throw new DbInsertErrorException();
                    orderStrings.Add(oneString);
                    CreatePeriodPosition(
                        oneString,
                        (int)hireTable.Rows[i].Cells[1].Value,
                        GetPositionPKByName(hireTable.Rows[i].Cells[3].Value.ToString()),
                        hireTable.Rows[i].Cells[7].Value.ToString());
                }
                if (MessageBox.Show("Приказ успешно добавлен!\nСохранить в Excel-файл?", "Успех", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    HireToExcel();
                HireTabReset();
                
            }
            catch (DbInsertErrorException ex)
            {
                string sql;
                foreach (int one in orderStrings)
                {
                    sql = "Delete from \"String_order\" where \"pk_string_order\" = " + one;
                    new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                }
                sql = "Delete from \"Order\" where \"pk_order\" = " + order;
                new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();
                return;
            }
            catch (EmptyTableError ETerr)
            {
                MessageBox.Show(ETerr.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private string GenereteOrderNum(int lenght) {
            string num = "";
            Random rand = new Random();
            for (int i = 0; i < lenght; i++) {
                num = num + rand.Next(0, 9);
            }
            return num;
        }

        private int CreateOrder(OrderType oType, string oNum, string date) {
            try {
                int newPK = 0;
                string oTypeName;
                orderDict.TryGetValue(oType, out oTypeName);
                string sql = "insert into \"Order\"(\"pk_type_order\", \"nomer\", \"data_order\")"
                    + " values ("
                    + "(select \"pk_type_order\" from \"TypeOrder\" where \"Name\" = '" + oTypeName + "'), "
                    + oNum + ", to_date('" + date + "', 'dd-MM-yyyy')) returning \"pk_order\"";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    newPK = reader.GetInt32(0);
                    reader.Close();
                }
                else {
                    reader.Close();
                    throw new Exception("Не удалось добавить приказ");
                }
                return newPK;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }

        private int CreateOrderString(int order, string contractNum, string contractDate, string reason) {
            try
            {
                int newPK = 0;
                string sql = "insert into \"String_order\"(\"pk_order\", \"Number_work_doc\",\"Work_doc_date\", \"Reason\")"
                    + " values ("
                    + order + ", '" + contractNum + "', to_date('" + contractDate + "', 'dd-MM-yyyy'), '" + reason + "') returning \"pk_string_order\"";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    newPK = (int)reader.GetValue(0);
                    reader.Close();
                }
                else
                {
                    reader.Close();
                    throw new Exception("Не удалось добавить строку приказа - приказ не добавлен");
                }
                return newPK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }

        private int CreatePeriodPosition(int stringOrder, int personalCard, int position, string startDate) {
            try
            {
                int newPK = 0;
                string sql = "insert into \"PeriodPosition\"(\"pk_position\", \"pk_personal_card\",\"pk_move_order\", \"DataFrom\")"
                    + " values ("
                    + position + ", " + personalCard + ", " + stringOrder + ", to_date('" + startDate + "', 'dd-MM-yyyy')) returning \"pk_period_position\"";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection.get_connect());
                NpgsqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    newPK = (int)reader.GetValue(0);
                    reader.Close();
                }
                else
                {
                    reader.Close();
                    throw new Exception("Не удалось добавить период");
                }
                return newPK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }

        private void ClosePeriodPosition(int person, string date) {
            try {
                    string sql = "Update \"PeriodPosition\" set \"DateTo\" = to_date('" + date + "', 'dd-MM-yyyy')"
                        + " where \"pk_personal_card\" = " + person + " and \"DateTo\" is null;";
                    new NpgsqlCommand(sql, connection.get_connect()).ExecuteNonQuery();

            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int GetPositionPKByName(string name) {

            string sql = "select \"pk_position\" from \"Position\""
                       + " where \"Name\" = '" + name + "'";
            NpgsqlDataReader reader = new NpgsqlCommand(sql, connection.get_connect()).ExecuteReader();
            reader.Read();
            int posPK = reader.GetInt32(0);
            reader.Close();
            return posPK;
        }

        private void HireToExcel()
        {
            Excel.Application app = new Excel.Application();
            string openFile = Environment.CurrentDirectory + "\\Orders\\Templates\\HireOrderTemplate.xls";
            app.Workbooks.Open(openFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing);
            app.DisplayAlerts = false;

            Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.Item[1];

            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            sheet.Range[sheet.Cells[9, "BW"], sheet.Cells[9, "CN"]] = hireDocNum.Text;
            sheet.Range[sheet.Cells[9, "CO"], sheet.Cells[9, "DG"]] = hireDocDate.Value.ToString(dateFormat);

            int currentRow = 18;
            for (int i = 0; i < hireTable.Rows.Count; i++, currentRow++)
            {
                sheet.Rows[currentRow].RowHeight = 30;

                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "W"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "W"]] = hireTable.Rows[i].Cells[0].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AF"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AF"]] = hireTable.Rows[i].Cells[1].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AG"], sheet.Cells[currentRow, "AS"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AG"], sheet.Cells[currentRow, "AS"]] = hireTable.Rows[i].Cells[2].Value.ToString();


                sheet.Range[sheet.Cells[currentRow, "AT"], sheet.Cells[currentRow, "BH"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AT"], sheet.Cells[currentRow, "BH"]] = hireTable.Rows[i].Cells[3].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "BI"], sheet.Cells[currentRow, "BS"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "BI"], sheet.Cells[currentRow, "BS"]] = hireTable.Rows[i].Cells[4].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "BT"], sheet.Cells[currentRow, "CA"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "BT"], sheet.Cells[currentRow, "CA"]] = hireTable.Rows[i].Cells[5].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "CB"], sheet.Cells[currentRow, "CI"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "CB"], sheet.Cells[currentRow, "CI"]] = hireTable.Rows[i].Cells[6].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "CJ"], sheet.Cells[currentRow, "CQ"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "CJ"], sheet.Cells[currentRow, "CQ"]] = hireTable.Rows[i].Cells[7].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "CR"], sheet.Cells[currentRow, "CX"]].Merge();

                sheet.Range[sheet.Cells[currentRow, "CY"], sheet.Cells[currentRow, "DH"]].Merge();

                sheet.Range[sheet.Cells[currentRow, "DI"], sheet.Cells[currentRow, "ED"]].Merge();

                
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "ED"]]).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "ED"]]).Cells.WrapText = true;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "ED"]]).Cells.Font.Size = 9;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "ED"]]).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "ED"]]).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
            currentRow++;
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "Z"]].Merge();
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "Z"]] = "Руководитель от организации";
            ((Excel.Range)(sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "Z"]])).Cells.Font.Bold = true;

            sheet.Range[sheet.Cells[currentRow, "AD"], sheet.Cells[currentRow, "BL"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "AD"], sheet.Cells[currentRow + 1, "BL"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "AD"], sheet.Cells[currentRow + 1, "BL"]] = "(должность)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "AD"], sheet.Cells[currentRow, "BL"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Range[sheet.Cells[currentRow, "BO"], sheet.Cells[currentRow, "CE"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "BO"], sheet.Cells[currentRow + 1, "CE"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "BO"], sheet.Cells[currentRow + 1, "CE"]] = "(личная подпись)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "BO"], sheet.Cells[currentRow, "CE"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Range[sheet.Cells[currentRow, "CJ"], sheet.Cells[currentRow, "ED"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "CJ"], sheet.Cells[currentRow + 1, "ED"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "CJ"], sheet.Cells[currentRow + 1, "ED"]] = "(расшифровка подписи)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "CJ"], sheet.Cells[currentRow, "ED"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            string fileDir = Environment.CurrentDirectory + "\\Orders";
            if (!Directory.Exists(fileDir))
                Directory.CreateDirectory(fileDir);

            string fileName = fileDir + "\\HIRE_" + hireDocNum.Text + "_" + hireDocDate.Value.ToString(dateFormat);
            if (File.Exists(fileName + ".xls"))
                fileName = fileName + "(" + DateTime.Now.ToString("dd-MM-yyyy HH-mm") + ")";
            
            app.Application.ActiveWorkbook.SaveAs(fileName + ".xls", Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, 
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            app.Application.ActiveWorkbook.Close();
            app.Quit();
            MessageBox.Show("Приказ сохранен по пути: " + fileName + ".xls", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void FireToExcel()
        {
            Excel.Application app = new Excel.Application();
            string openFile = Environment.CurrentDirectory + "\\Orders\\Templates\\FireOrderTemplate.xls";
            app.Workbooks.Open(openFile, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing);
            app.DisplayAlerts = false;

            Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.Item[1];

            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            
            
            sheet.Range[sheet.Cells[9, "AA"], sheet.Cells[9, "AF"]] = fireDocNum.Text;

            sheet.Range[sheet.Cells[9, "AG"], sheet.Cells[9, "AN"]] = fireDocDate.Value.ToString(dateFormat);
            int currentRow = 18;
            for (int i = 0; i < fireTable.Rows.Count; i++, currentRow++)
            {
                sheet.Rows[currentRow].RowHeight = 30;

                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "G"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "G"]] = fireTable.Rows[i].Cells[0].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "H"], sheet.Cells[currentRow, "J"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "H"], sheet.Cells[currentRow, "J"]] = fireTable.Rows[i].Cells[1].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "O"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "O"]] = fireTable.Rows[i].Cells[2].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "T"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "T"]] = fireTable.Rows[i].Cells[3].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "U"], sheet.Cells[currentRow, "X"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "U"], sheet.Cells[currentRow, "X"]] = fireTable.Rows[i].Cells[4].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "Y"], sheet.Cells[currentRow, "AA"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "Y"], sheet.Cells[currentRow, "AA"]] = fireTable.Rows[i].Cells[5].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]] = fireTable.Rows[i].Cells[7].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AJ"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AJ"]] = fireTable.Rows[i].Cells[6].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AK"], sheet.Cells[currentRow, "AP"]].Merge();

                sheet.Range[sheet.Cells[currentRow, "AQ"], sheet.Cells[currentRow, "AV"]].Merge();

                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.WrapText = true;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.Font.Size = 9;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
            currentRow++;
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "J"]].Merge();
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "J"]] = "Руководитель организации";
            ((Excel.Range)(sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "J"]])).Cells.Font.Size = 12;

            sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "T"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "K"], sheet.Cells[currentRow + 1, "T"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "K"], sheet.Cells[currentRow + 1, "T"]] = "(должность)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "T"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Range[sheet.Cells[currentRow, "V"], sheet.Cells[currentRow, "AC"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "V"], sheet.Cells[currentRow + 1, "AC"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "V"], sheet.Cells[currentRow + 1, "AC"]] = "(личная подпись)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "V"], sheet.Cells[currentRow, "AC"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Range[sheet.Cells[currentRow, "AE"], sheet.Cells[currentRow, "AQ"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "AE"], sheet.Cells[currentRow + 1, "AQ"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "AE"], sheet.Cells[currentRow + 1, "AQ"]] = "(расшифровка подписи)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "AE"], sheet.Cells[currentRow, "AQ"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            currentRow += 2;

            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "N"]].Merge();
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "N"]] = "Мотивированное мнение выборного профсоюзного органа в письменной форме";
            ((Excel.Range)(sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]])).RowHeight = 25;
            ((Excel.Range)(sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]])).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            currentRow++;
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]].Merge();
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]] = "(от \"____\" ____________ 20__ г.  № _____________ ) рассмотренно";
            ((Excel.Range)(sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]])).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            string fileDir = Environment.CurrentDirectory + "\\Orders";
            if (!Directory.Exists(fileDir))
                Directory.CreateDirectory(fileDir);
            string fileName = fileDir + "\\FIRE_" + hireDocNum.Text + "_" + hireDocDate.Value.ToString(dateFormat);
            if (File.Exists(fileName + ".xls"))
                fileName = fileName + "(" + DateTime.Now.ToString("dd-MM-yyyy HH-mm") + ")";

            app.Application.ActiveWorkbook.SaveAs(fileName + ".xls", Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            app.Application.ActiveWorkbook.Close();
            app.Quit();

            MessageBox.Show("Приказ сохранен по пути: " + fileName + ".xls", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void MoveToExcel()
        {
            Excel.Application app = new Excel.Application();
            string openFile = Environment.CurrentDirectory + "\\Orders\\Templates\\MoveOrderTemplate.xls";
            app.Workbooks.Open(openFile, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing);
            app.DisplayAlerts = false;

            Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.Item[1];

            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            sheet.Range[sheet.Cells[9, "AA"], sheet.Cells[9, "AF"]] = moveDocNum.Text;

            sheet.Range[sheet.Cells[9, "AG"], sheet.Cells[9, "AL"]] = moveDocDate.Value.ToString(dateFormat);

            int currentRow = 17;
            for (int i = 0; i < moveTable.Rows.Count; i++, currentRow++)
            {
                sheet.Rows[currentRow].RowHeight = 30;

                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "H"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "H"]] = moveTable.Rows[i].Cells[0].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "I"], sheet.Cells[currentRow, "K"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "I"], sheet.Cells[currentRow, "K"]] = moveTable.Rows[i].Cells[1].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "L"], sheet.Cells[currentRow, "O"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "L"], sheet.Cells[currentRow, "O"]] = moveTable.Rows[i].Cells[2].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "S"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "S"]] = moveTable.Rows[i].Cells[3].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "T"], sheet.Cells[currentRow, "W"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "T"], sheet.Cells[currentRow, "W"]] = moveTable.Rows[i].Cells[4].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AA"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AA"]] = moveTable.Rows[i].Cells[5].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]] = moveTable.Rows[i].Cells[6].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AH"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AH"]] = moveTable.Rows[i].Cells[9].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AI"], sheet.Cells[currentRow, "AK"]].Merge();

                sheet.Range[sheet.Cells[currentRow, "AL"], sheet.Cells[currentRow, "AN"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AL"], sheet.Cells[currentRow, "AN"]] = moveTable.Rows[i].Cells[7].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AO"], sheet.Cells[currentRow, "AQ"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AO"], sheet.Cells[currentRow, "AQ"]] = moveTable.Rows[i].Cells[8].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AR"], sheet.Cells[currentRow, "AV"]].Merge();

                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.WrapText = true;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.Font.Size = 9;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "AV"]]).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            }
            currentRow++;
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "J"]].Merge();
            sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "J"]] = "Руководитель организации";
            ((Excel.Range)(sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "J"]])).Cells.Font.Size = 12;

            sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "T"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "K"], sheet.Cells[currentRow + 1, "T"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "K"], sheet.Cells[currentRow + 1, "T"]] = "(должность)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "T"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Range[sheet.Cells[currentRow, "V"], sheet.Cells[currentRow, "AC"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "V"], sheet.Cells[currentRow + 1, "AC"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "V"], sheet.Cells[currentRow + 1, "AC"]] = "(личная подпись)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "V"], sheet.Cells[currentRow, "AC"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            sheet.Range[sheet.Cells[currentRow, "AE"], sheet.Cells[currentRow, "AQ"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "AE"], sheet.Cells[currentRow + 1, "AQ"]].Merge();
            sheet.Range[sheet.Cells[currentRow + 1, "AE"], sheet.Cells[currentRow + 1, "AQ"]] = "(расшифровка подписи)";
            ((Excel.Range)sheet.Range[sheet.Cells[currentRow, "AE"], sheet.Cells[currentRow, "AQ"]]).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            string fileDir = Environment.CurrentDirectory + "\\Orders";
            if (!Directory.Exists(fileDir))
                Directory.CreateDirectory(fileDir);
            string fileName = fileDir + "\\MOVE_" + hireDocNum.Text + "_" + hireDocDate.Value.ToString(dateFormat);
            if (File.Exists(fileName + ".xls"))
                fileName = fileName + "(" + DateTime.Now.ToString("dd-MM-yyyy HH-mm") + ")";

            app.Application.ActiveWorkbook.SaveAs(fileName + ".xls", Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            app.Application.ActiveWorkbook.Close();
            app.Quit();

            MessageBox.Show("Приказ сохранен по пути: " + fileName + ".xls", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
