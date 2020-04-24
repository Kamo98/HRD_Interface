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
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelDeptApp1
{
	public partial class StaffTurnover : Form
	{
		BindingList<Department> departments = new BindingList<Department>();
		BindingList<Occupation> occupations = new BindingList<Occupation>();

		Dictionary<string, int> str2idTO = new Dictionary<string, int>();

		public StaffTurnover()
		{
			InitializeComponent();
			get_type_orders();

			FillDepts();
		}

		private void btn_buildGraphic_Click(object sender, EventArgs e)
		{

			//Настройка графика
			chart1.Series.Clear();
			var s1 = new Series();
			var s2 = new Series();
			s1.ChartType = s2.ChartType = SeriesChartType.Line;
			chart1.Series.Add(s1);
			chart1.Series.Add(s2);

			chart1.Series[0].XValueType = ChartValueType.DateTime;
			chart1.Series[1].XValueType = ChartValueType.DateTime;
			chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM";
			chart1.ChartAreas[0].AxisX.Interval = 1;
			chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
			chart1.ChartAreas[0].AxisX.IntervalOffset = 1;

			chart1.Series[0].XValueType = ChartValueType.DateTime;
			chart1.Series[1].XValueType = ChartValueType.DateTime;
			chart1.Series[0].Name = "Принятых";
			chart1.Series[1].Name = "Уволенных";
			chart1.Series[0].ChartType = SeriesChartType.Spline;
			chart1.Series[1].ChartType = SeriesChartType.Spline;
			chart1.Series[0].BorderWidth = 3;
			chart1.Series[1].BorderWidth = 3;
			
			DateTime dateStart = new DateTime((int)yearFrom.Value, (int)monthFrom.Value, 1);
			DateTime dateFinish = new DateTime((int)yearTo.Value, (int)monthTo.Value, DateTime.DaysInMonth((int)yearTo.Value, (int)monthTo.Value));
			
			chart1.ChartAreas[0].AxisX.Minimum = dateStart.ToOADate();
			chart1.ChartAreas[0].AxisX.Maximum = dateFinish.ToOADate();


			//ПРойти по всем месяцам
			for (DateTime dateFrom = dateStart; dateFrom < dateFinish; dateFrom = next_month(dateFrom))
			{
				DateTime dateTo = new DateTime(dateFrom.Year, dateFrom.Month, DateTime.DaysInMonth(dateFrom.Year, dateFrom.Month));

				string sqlIn = "";		//Поиск приказов на приём
				string sqlOut = "";		//Поиск приказов на увольнение

				//По всей организации
				if (cb_allUnits.Checked)
				{
					sqlIn = "select count(s.\"pk_string_order\")" +
					" from \"Order\" o, \"String_order\" s" +
						" where o.\"pk_type_order\" = '" +
						str2idTO["Приём"] + "' and s.\"pk_order\" = o.\"pk_order\" " +
						" and o.\"data_order\" >= '" + date2str(dateFrom) + 
						"' and o.\"data_order\" <= '" + date2str(dateTo) + "'";

					sqlOut = "select count(s.\"pk_string_order\")" +
					" from \"Order\" o, \"String_order\" s" +
						" where o.\"pk_type_order\" = '" +
						str2idTO["Увольнение"] + "' and s.\"pk_order\" = o.\"pk_order\" " +
					   " and o.\"data_order\" >= '" + date2str(dateFrom) +
						"' and o.\"data_order\" <= '" + date2str(dateTo) + "'";

				} else if (cb_allPositions.Checked)			//По подразделению
				{
					int idUnit = (department.SelectedItem as Department).Id;

					sqlIn = "select count(s.\"pk_string_order\")" +
					" from \"Order\" o, \"String_order\" s, \"PeriodPosition\" p, \"Position\" d" +
						" where s.\"pk_order\" = o.\"pk_order\" " +
						" and p.\"pk_move_order\" = s.\"pk_string_order\" " +
						" and p.\"pk_position\" = d.\"pk_position\" " +
						" and o.\"data_order\" >= '" + date2str(dateFrom) +
						"' and o.\"data_order\" <= '" + date2str(dateTo) + "'" +
						" and (o.\"pk_type_order\" = '" + str2idTO["Приём"] + "' " +
						" or o.\"pk_type_order\" = '" + str2idTO["Перевод"] + "') " +
						" and d.\"pk_unit\" = '" + idUnit + "' ";

					sqlOut = "select count(s.\"pk_string_order\")" +
					" from \"Order\" o, \"String_order\" s, \"PeriodPosition\" p, \"Position\" d" +
						" where s.\"pk_order\" = o.\"pk_order\" " +
						" and p.\"pk_position\" = d.\"pk_position\" " +
						" and o.\"data_order\" >= '" + date2str(dateFrom) +
						"' and o.\"data_order\" <= '" + date2str(dateTo) + "'" +
						" and p.\"pk_fire_order_string\" = s.\"pk_string_order\" " +
						" and o.\"pk_type_order\" = '" + str2idTO["Увольнение"] + "' " +
						" and d.\"pk_unit\" = '" + idUnit + "' ";

				} else
				{
					int idPosition = (position.SelectedItem as Occupation).Id;

					sqlIn = "select count(s.\"pk_string_order\")" +
					" from \"Order\" o, \"String_order\" s, \"PeriodPosition\" p" +
						" where s.\"pk_order\" = o.\"pk_order\" " +
						" and p.\"pk_move_order\" = s.\"pk_string_order\" " +
						" and o.\"data_order\" >= '" + date2str(dateFrom) +
						"' and o.\"data_order\" <= '" + date2str(dateTo) + "'" +
						" and (o.\"pk_type_order\" = '" + str2idTO["Приём"] + "' " +
						" or o.\"pk_type_order\" = '" + str2idTO["Перевод"] + "') " +
						" and p.\"pk_position\" = '" + idPosition + "' ";

					sqlOut = "select count(s.\"pk_string_order\")" +
					" from \"Order\" o, \"String_order\" s, \"PeriodPosition\" p" +
						" where s.\"pk_order\" = o.\"pk_order\" " +
						" and o.\"data_order\" >= '" + date2str(dateFrom) +
						"' and o.\"data_order\" <= '" + date2str(dateTo) + "'" +
						" and p.\"pk_fire_order_string\" = s.\"pk_string_order\" " +
						" and o.\"pk_type_order\" = '" + str2idTO["Увольнение"] + "' " +
						" and p.\"pk_position\" = '" + idPosition + "' ";
				}



				//Выполнение запросов и добавление результатов в серии графика
				long countIn = get_count_orders(sqlIn);
				long countOut = get_count_orders(sqlOut);

				if (rb_accumulative.Checked && chart1.Series[0].Points.Count != 0)
				{
					countIn += (long)chart1.Series[0].Points[chart1.Series[0].Points.Count - 1].YValues[0];
					countOut += (long)chart1.Series[1].Points[chart1.Series[1].Points.Count - 1].YValues[0];
				}

				chart1.Series[0].Points.AddXY(dateFrom, countIn);
				chart1.Series[1].Points.AddXY(dateFrom, countOut);
			}
			
		}


		private string date2str(DateTime d)
		{
			return d.Year + "-" + d.Month + "-" + d.Day;
		}

		private DateTime next_month(DateTime date)
		{
			DateTime dateRes;
			if (date.Month == 12)
				dateRes = new DateTime(date.Year + 1, 1, 1);          //Первый день месяца
			else
				dateRes = new DateTime(date.Year, date.Month + 1, 1); //Первый день месяца
			return dateRes;
		}


		private void get_type_orders ()
		{
			string sql = "select * from \"TypeOrder\"";
			try
			{
				if (Connection.get_connect() == null)
					throw new NullReferenceException("Не удалось подключиться к базе данных");
				NpgsqlCommand command = new NpgsqlCommand(sql, Connection.get_connect());
				NpgsqlDataReader reader = command.ExecuteReader();
				foreach (DbDataRecord record in reader)
				{
					str2idTO.Add((string)record["Name"], (int)record["pk_type_order"]);
				}
				reader.Close();
				
			}
			catch (NullReferenceException ex)
			{
				MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private long get_count_orders(string sql)
		{
			long res = 0;
			try
			{
				if (Connection.get_connect() == null)
					throw new NullReferenceException("Не удалось подключиться к базе данных");
				NpgsqlCommand command = new NpgsqlCommand(sql, Connection.get_connect());
				NpgsqlDataReader reader = command.ExecuteReader();

				foreach (DbDataRecord record in reader)
				{
					object[] obj = new object[record.FieldCount];
					record.GetValues(obj);
					
					res = (long)(obj[0]);
					break;
				}
				reader.Close();
			}
			catch (NullReferenceException ex)
			{
				MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			return res;
		}

		private void FillDepts()
		{
			try
			{
				string sql = "select * from \"Unit\";";
				if (Connection.get_connect() == null)
					throw new NullReferenceException("Не удалось подключиться к базе данных");
				NpgsqlCommand command = new NpgsqlCommand(sql, Connection.get_connect());
				NpgsqlDataReader reader = command.ExecuteReader();
				foreach (DbDataRecord record in reader)
				{
					Department newDept = new Department((int)record["pk_unit"], (string)record["Name"]);
					departments.Add(newDept);

				}
				reader.Close();

				department.DataSource = departments;
				department.DisplayMember = "Name";
			}
			catch (NullReferenceException ex)
			{
				MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
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
				NpgsqlCommand command = new NpgsqlCommand(sql, Connection.get_connect());
				NpgsqlDataReader reader = command.ExecuteReader();
				foreach (DbDataRecord record in reader)
				{
					occupations.Add(new Occupation((int)record["pk_position"], (string)record["Name"], (decimal)record["Rate"]));
				}
				reader.Close();

				position.DataSource = occupations;
				position.DisplayMember  = "Name";
				
			}
			catch (Exception ex)
			{
				MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

		}

		private void StaffTurnover_FormClosed(object sender, FormClosedEventArgs e)
		{
			Form ifrm = Application.OpenForms[Application.OpenForms.Count - 1];
			ifrm.Show();
		}

		private void cb_allUnits_CheckedChanged(object sender, EventArgs e)
		{
			if (cb_allUnits.Checked)
			{
				department.Enabled = false;
				department.Text = "";
				cb_allPositions.Checked = false;
				cb_allPositions.Enabled = false;
				position.Enabled = false;
				position.Text = "";

			} else
			{
				department.Enabled = true;
				department.SelectedIndex = 0;
				cb_allPositions.Checked = false;
				cb_allPositions.Enabled = true;
				position.Enabled = true;
			}
		}

		private void cb_allPositions_CheckedChanged(object sender, EventArgs e)
		{
			if (cb_allPositions.Checked)
			{
				position.Enabled = false;
				position.Text = "";
			}
			else
			{
				position.Enabled = true;
				position.SelectedIndex = 0;
			}
		}
	}
}
