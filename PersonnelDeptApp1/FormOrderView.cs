using System;
using System.Data.Common;
using System.Windows.Forms;

namespace PersonnelDeptApp1
{
    public partial class FormOrderView : Form
    {
        int orderId;
        string orderType;
        string num;
        DateTime date;
        Npgsql.NpgsqlConnection connection = Connection.get_connect();
        
        public FormOrderView(int id, string orderType)
        {
            InitializeComponent();
            orderId = id;
            this.orderType = orderType;
        }

        private void FormOrderView_Load(object sender, EventArgs e)
        {
            string sql = "select \"nomer\", \"data_order\" from \"Order\" where \"pk_order\" = " + orderId;
            Npgsql.NpgsqlDataReader reader = new Npgsql.NpgsqlCommand(sql, connection).ExecuteReader();
            reader.Read();
            num = reader.GetString(0);
            date = reader.GetDateTime(1);
            reader.Close();
            label1.Text = "Договор №" + num + " от " + date.ToString("dd-MM-yyyy");
            switch (orderType) {
                case "Приём":
                    LoadHireOrder();
                    break;
                case "Увольнение":
                    LoadFireOrder();
                    break;
                case "Перевод":
                    LoadMoveOrder();
                    break;
            }
        }

        private void LoadHireOrder() 
        {
            orderTable.Columns.Add("FIO", "Сотрудник");
            orderTable.Columns.Add("ID", "Номер личной карточки");
            orderTable.Columns.Add("Dep", "Подразделение");
            orderTable.Columns.Add("Pos", "Должность");
            orderTable.Columns.Add("Salary", "Тарифная ставка");
            orderTable.Columns.Add("ContractNum", "Номер договора");
            orderTable.Columns.Add("ContractDate", "Дата создания договора");
            orderTable.Columns.Add("StartWork", "Период работы с (дата)");

            orderTable.Columns[0].FillWeight = 16;
            orderTable.Columns[1].FillWeight = 7;
            orderTable.Columns[2].FillWeight = 20;
            orderTable.Columns[3].FillWeight = 20;
            orderTable.Columns[4].FillWeight = 7;
            orderTable.Columns[5].FillWeight = 10;
            orderTable.Columns[6].FillWeight = 10;
            orderTable.Columns[7].FillWeight = 10;

            string sql = "select * from get_one_hire_order(" + orderId + ")";
            Npgsql.NpgsqlDataReader reader = new Npgsql.NpgsqlCommand(sql, connection).ExecuteReader();
            foreach (DbDataRecord record in reader) {
                orderTable.Rows.Add(
                    record[0],
                    record[1],
                    record[2],
                    record[3],
                    record[4],
                    record[5],
                    record[6],
                    (record.GetDateTime(7)).ToShortDateString());
            }
            reader.Close();
        }
        private void LoadFireOrder()
        {
            orderTable.Columns.Add("FIO", "Сотрудник");
            orderTable.Columns.Add("ID", "Номер личной карточки");
            orderTable.Columns.Add("Dep", "Подразделение");
            orderTable.Columns.Add("Pos", "Должность");
            orderTable.Columns.Add("ContractNum", "Номер договора");
            orderTable.Columns.Add("ContractDate", "Дата создания договора");
            orderTable.Columns.Add("Reason", "Основание");
            orderTable.Columns.Add("StartWork", "Период работы по (дата)");

            orderTable.Columns[0].FillWeight = 15;
            orderTable.Columns[1].FillWeight = 5;
            orderTable.Columns[2].FillWeight = 18;
            orderTable.Columns[3].FillWeight = 18;
            orderTable.Columns[4].FillWeight = 9;
            orderTable.Columns[5].FillWeight = 10;
            orderTable.Columns[6].FillWeight = 15;
            orderTable.Columns[7].FillWeight = 10;

            string sql = "select * from get_one_fire_order(" + orderId + ")";
            Npgsql.NpgsqlDataReader reader = new Npgsql.NpgsqlCommand(sql, connection).ExecuteReader();
            foreach (DbDataRecord record in reader)
            {
                orderTable.Rows.Add(
                    record[0],
                    record[1],
                    record[2],
                    record[3],
                    record[4],
                    record[5],
                    record[6],
                    (record.GetDateTime(7)).ToShortDateString());
            }
            reader.Close();
        }
        private void LoadMoveOrder()
        {
            orderTable.Columns.Add("FIO", "Сотрудник");
            orderTable.Columns.Add("ID", "Номер личной карточки");
            orderTable.Columns.Add("Dep", "Структурное подразделение (прежнее)");
            orderTable.Columns.Add("DepNew", "Структурное подразделение (новое)");
            orderTable.Columns.Add("Pos", "Должность (прежняя)");
            orderTable.Columns.Add("PosNew", "Должность (новая)");
            orderTable.Columns.Add("Reason", "Тарифная ставка");
            orderTable.Columns.Add("ContractNum", "Трудовой договор: номер");
            orderTable.Columns.Add("ContractDate", "Трудовой договор: дата");
            orderTable.Columns.Add("StartWork", "Период работы по (дата)");

            orderTable.Columns[0].FillWeight = 15;
            orderTable.Columns[1].FillWeight = 5;
            orderTable.Columns[2].FillWeight = 13;
            orderTable.Columns[3].FillWeight = 13;
            orderTable.Columns[4].FillWeight = 13;
            orderTable.Columns[5].FillWeight = 13;
            orderTable.Columns[6].FillWeight = 5;
            orderTable.Columns[7].FillWeight = 5;
            orderTable.Columns[8].FillWeight = 5;
            orderTable.Columns[9].FillWeight = 5;

            string sql = "select * from get_one_move_order(" + orderId + ")";
            Npgsql.NpgsqlDataReader reader = new Npgsql.NpgsqlCommand(sql, connection).ExecuteReader();
            foreach (DbDataRecord record in reader)
            {
                orderTable.Rows.Add(
                    record[0],
                    record[1],
                    record[2],
                    record[4],
                    record[3],
                    record[5],
                    record[6],
                    record[7],
                    record[8],
                    (record.GetDateTime(9)).ToShortDateString());
            }
            reader.Close();

        }
    }
}
