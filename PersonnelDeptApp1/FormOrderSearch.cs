using System;
using Npgsql;
using System.Windows.Forms;
using System.ComponentModel;
using System.Data.Common;

namespace PersonnelDeptApp1
{
    public partial class FormOrderSearch : Form
    {
        NpgsqlConnection connection = Connection.get_connect();
        BindingList<Employee> employees = new BindingList<Employee>();
        BindingList<OrderType> orderTypes = new BindingList<OrderType>();
        public FormOrderSearch()
        {
            InitializeComponent();
        }



        private void FormOrderSearch_Load(object sender, EventArgs e)
        {
            string sql = "Select * from \"TypeOrder\"";
            NpgsqlDataReader reader = new NpgsqlCommand(sql, connection).ExecuteReader();
            foreach (DbDataRecord record in reader) {
                orderTypes.Add(new OrderType((int)record["pk_type_order"], record["Name"].ToString()));
            }
            reader.Close();

            sql = "Select * from getlist_by_substring('')";
            reader = new NpgsqlCommand(sql, connection).ExecuteReader();
            foreach (DbDataRecord record in reader)
            {
                employees.Add(new Employee((int)record[1], record[0].ToString()));
            }
            reader.Close();

            orderTypesCB.DataSource = orderTypes;
            orderTypesCB.DisplayMember = "Name";

            empListCB.DataSource = employees;
            empListCB.DisplayMember = "FIO";
            empListCB.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            empListCB.AutoCompleteSource = AutoCompleteSource.ListItems;
            
        }

        private void empListCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedEmpId.Text = ((sender as ComboBox).SelectedItem as Employee).Id.ToString();
        }

        private void FindBTN_Click(object sender, EventArgs e)
        {
            ordersTable.Rows.Clear();
            string sql = "";
            switch ((orderTypesCB.SelectedItem as OrderType).Name){
                case "Приём":
                    PositionCol.HeaderText = "Принят на должность";
                    sql = "select * from get_hire_orders(" + selectedEmpId.Text + ") order by \"orderdate\"";
                    break;
                case "Увольнение":
                    PositionCol.HeaderText = "Уволен с должности";
                    sql = "select * from get_fire_orders(" + selectedEmpId.Text + ") order by \"orderdate\"";
                    break;
                case "Перевод":
                    PositionCol.HeaderText = "Переведен на должность";
                    sql = "select * from get_move_orders(" + selectedEmpId.Text + ") order by \"orderdate\"";
                    break;
            }
            NpgsqlDataReader reader = new NpgsqlCommand(sql, connection).ExecuteReader();
            foreach (DbDataRecord record in reader) {
                ordersTable.Rows.Add(
                    (int)record["orderid"],
                    record["orderNum"],
                    record["orderdate"],
                    record["pos"],
                    record["contrnum"],
                    "Показать приказ"
                    );
            }
            reader.Close();
        }

        private void FormOrderSearch_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.OpenForms[Application.OpenForms.Count - 1].Show();
        }

        private void ordersTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != ShowOrderBtnCol.Index)
                return;
            DateTimeConverter converter = new DateTimeConverter();

            new FormOrderView((int)ordersTable.Rows[e.RowIndex].Cells[0].Value, (orderTypesCB.SelectedItem as OrderType).Name).ShowDialog();
        }
    }
}
