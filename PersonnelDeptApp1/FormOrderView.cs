using System;
using System.Data.Common;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace PersonnelDeptApp1
{
    public partial class FormOrderView : Form
    {
        int orderId;
        string orderType;
        string num;
        DateTime date;
        const string dateFormat = "dd-MM-yyyy";
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
            label1.Text = "Приказ №" + num + " от " + date.ToString("dd-MM-yyyy");
            switch (orderType)
            {
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

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            switch (orderType)
            {
                case "Приём":
                    HireToExcel();
                    break;
                case "Увольнение":
                    FireToExcel();
                    break;
                case "Перевод":
                    MoveToExcel();
                    break;
            }
            button1.Enabled = true;
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
            sheet.Range[sheet.Cells[9, "BW"], sheet.Cells[9, "CN"]] = num;
            sheet.Range[sheet.Cells[9, "CO"], sheet.Cells[9, "DG"]] = date.ToString(dateFormat);

            int currentRow = 18;
            for (int i = 0; i < orderTable.Rows.Count; i++, currentRow++)
            {
                sheet.Rows[currentRow].RowHeight = 30;

                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "W"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "W"]] = orderTable.Rows[i].Cells[0].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AF"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AF"]] = orderTable.Rows[i].Cells[1].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AG"], sheet.Cells[currentRow, "AS"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AG"], sheet.Cells[currentRow, "AS"]] = orderTable.Rows[i].Cells[2].Value.ToString();


                sheet.Range[sheet.Cells[currentRow, "AT"], sheet.Cells[currentRow, "BH"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AT"], sheet.Cells[currentRow, "BH"]] = orderTable.Rows[i].Cells[3].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "BI"], sheet.Cells[currentRow, "BS"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "BI"], sheet.Cells[currentRow, "BS"]] = orderTable.Rows[i].Cells[4].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "BT"], sheet.Cells[currentRow, "CA"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "BT"], sheet.Cells[currentRow, "CA"]] = orderTable.Rows[i].Cells[5].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "CB"], sheet.Cells[currentRow, "CI"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "CB"], sheet.Cells[currentRow, "CI"]] = orderTable.Rows[i].Cells[6].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "CJ"], sheet.Cells[currentRow, "CQ"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "CJ"], sheet.Cells[currentRow, "CQ"]] = orderTable.Rows[i].Cells[7].Value.ToString();

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

            string fileName = fileDir + "\\HIRE_" + num + "_" + date.ToString(dateFormat);
            if (File.Exists(fileName + ".xls"))
                fileName = fileName + "(" + DateTime.Now.ToString("dd-MM-yyyy HH-mm") + ")";

            app.Application.ActiveWorkbook.SaveAs(fileName + ".xls", Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (MessageBox.Show("Приказ сохранен по пути: " + fileName + ".xls\nПоказать в Excel?", "Успех", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                app.Visible = true;
            else
            {
                app.Application.ActiveWorkbook.Close();
                app.Quit();
            }
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



            sheet.Range[sheet.Cells[9, "AA"], sheet.Cells[9, "AF"]] = num;

            sheet.Range[sheet.Cells[9, "AG"], sheet.Cells[9, "AN"]] = date.ToString(dateFormat);
            int currentRow = 18;
            for (int i = 0; i < orderTable.Rows.Count; i++, currentRow++)
            {
                sheet.Rows[currentRow].RowHeight = 30;

                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "G"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "G"]] = orderTable.Rows[i].Cells[0].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "H"], sheet.Cells[currentRow, "J"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "H"], sheet.Cells[currentRow, "J"]] = orderTable.Rows[i].Cells[1].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "O"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "K"], sheet.Cells[currentRow, "O"]] = orderTable.Rows[i].Cells[2].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "T"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "T"]] = orderTable.Rows[i].Cells[3].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "U"], sheet.Cells[currentRow, "X"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "U"], sheet.Cells[currentRow, "X"]] = orderTable.Rows[i].Cells[4].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "Y"], sheet.Cells[currentRow, "AA"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "Y"], sheet.Cells[currentRow, "AA"]] = orderTable.Rows[i].Cells[5].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]] = orderTable.Rows[i].Cells[7].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AJ"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AJ"]] = orderTable.Rows[i].Cells[6].Value.ToString();

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
            string fileName = fileDir + "\\FIRE_" + num + "_" + date.ToString(dateFormat);
            if (File.Exists(fileName + ".xls"))
                fileName = fileName + "(" + DateTime.Now.ToString("dd-MM-yyyy HH-mm") + ")";

            app.Application.ActiveWorkbook.SaveAs(fileName + ".xls", Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (MessageBox.Show("Приказ сохранен по пути: " + fileName + ".xls\nПоказать в Excel?", "Успех", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                app.Visible = true;
            else
            {
                app.Application.ActiveWorkbook.Close();
                app.Quit();
            }
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

            sheet.Range[sheet.Cells[8, "AA"], sheet.Cells[8, "AF"]] = num;

            sheet.Range[sheet.Cells[8, "AG"], sheet.Cells[8, "AL"]] = date.ToString(dateFormat);

            int currentRow = 17;
            for (int i = 0; i < orderTable.Rows.Count; i++, currentRow++)
            {
                sheet.Rows[currentRow].RowHeight = 30;

                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "H"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "A"], sheet.Cells[currentRow, "H"]] = orderTable.Rows[i].Cells[0].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "I"], sheet.Cells[currentRow, "K"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "I"], sheet.Cells[currentRow, "K"]] = orderTable.Rows[i].Cells[1].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "L"], sheet.Cells[currentRow, "O"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "L"], sheet.Cells[currentRow, "O"]] = orderTable.Rows[i].Cells[2].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "S"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "P"], sheet.Cells[currentRow, "S"]] = orderTable.Rows[i].Cells[3].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "T"], sheet.Cells[currentRow, "W"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "T"], sheet.Cells[currentRow, "W"]] = orderTable.Rows[i].Cells[4].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AA"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "X"], sheet.Cells[currentRow, "AA"]] = orderTable.Rows[i].Cells[5].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AB"], sheet.Cells[currentRow, "AE"]] = orderTable.Rows[i].Cells[6].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AH"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AF"], sheet.Cells[currentRow, "AH"]] = orderTable.Rows[i].Cells[9].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AI"], sheet.Cells[currentRow, "AK"]].Merge();

                sheet.Range[sheet.Cells[currentRow, "AL"], sheet.Cells[currentRow, "AN"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AL"], sheet.Cells[currentRow, "AN"]] = orderTable.Rows[i].Cells[7].Value.ToString();

                sheet.Range[sheet.Cells[currentRow, "AO"], sheet.Cells[currentRow, "AQ"]].Merge();
                sheet.Range[sheet.Cells[currentRow, "AO"], sheet.Cells[currentRow, "AQ"]] = orderTable.Rows[i].Cells[8].Value.ToString();

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
            string fileName = fileDir + "\\MOVE_" + num + "_" + date.ToString(dateFormat);
            if (File.Exists(fileName + ".xls"))
                fileName = fileName + "(" + DateTime.Now.ToString("dd-MM-yyyy HH-mm") + ")";

            app.Application.ActiveWorkbook.SaveAs(fileName + ".xls", Type.Missing, Type.Missing, Type.Missing,
                          Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (MessageBox.Show("Приказ сохранен по пути: " + fileName + ".xls\nПоказать в Excel?", "Успех", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                app.Visible = true;
            else
            {
                app.Application.ActiveWorkbook.Close();
                app.Quit();
            }
        }
    }
}
