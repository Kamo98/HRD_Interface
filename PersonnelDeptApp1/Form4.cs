using HRD_GenerateData;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        ListBox employeesVars = new ListBox();
        BindingList<Employee> employees = new BindingList<Employee>();
        BindingList<Department> departments = new BindingList<Department>();
        BindingList<Occupation> occupations = new BindingList<Occupation>();
        Connection connection = Connection.get_instance("postgres", "Ntcnbhjdfybt_01");
        public Form4()
        {
            InitializeComponent();
            FillDepts();

            employeesVars.DataSource = employees;
            employeesVars.DisplayMember = "FIO";
            employeesVars.ValueMember = "Id";

            

            
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
            employeesVars.Location = new Point(
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
                hireTable.Rows.Add(
                    (employeesVars.SelectedItem as Employee).FIO,
                    (employeesVars.SelectedItem as Employee).Id,
                    (hireDepartment.SelectedItem as Department).Name,
                    (hireOccup.SelectedItem as Occupation).Name,
                    (hireOccup.SelectedItem as Occupation).Tarif,
                    hireContractNum.Text,
                    hireContractDate.Value.ToString("dd-MM-yyyy"),
                    startWork.Value.ToString("dd-MM-yyyy")
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
            HireTabReset();
            MoveTabReset();
            FireTabReset();
        }

        private void HireTabReset() {
            hireFIO.Text = "";
            employees.Clear();
            hireContractNum.Text = "";
            hireDocNum.Text = "";
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
            moveDocNum.Text = "";
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
            fireDocNum.Text = "";
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
                fireTable.Rows.Add(
                    (employeesVars.SelectedItem as Employee).FIO,
                    (employeesVars.SelectedItem as Employee).Id,
                    fireDepartment.Text,
                    fireOccup.Text,
                    fireContractNum.Text,
                    fireContractDate.Value.ToString("dd-MM-yyyy"),
                    fireReason.Text,
                    endWork.Value.ToString("dd-MM-yyyy")
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
                moveTable.Rows.Add(
                    (employeesVars.SelectedItem as Employee).FIO,
                    (employeesVars.SelectedItem as Employee).Id,
                    moveDepartmentOld.Text,
                    (moveDepartmentNew.SelectedItem as Department).Name,
                    moveOccupOld.Text,
                    (moveOccupNew.SelectedItem as Occupation).Name,
                    moveTarif.Text,
                    moveContractNum.Text,
                    moveContractDate.Value.ToString("dd-MM-yyyy"),
                    moveContractDate.Value.ToString("dd-MM-yyyy")
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
            MoveTabReset();
        }

        private void addFireOrder_Click(object sender, EventArgs e)
        {
            FireTabReset();
        }

        private void moveHireOrder_Click(object sender, EventArgs e)
        {
            HireTabReset();
        }
    }
}
