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
	public partial class StatTimeTracking : Form
	{
		BindingList<Department> departments = new BindingList<Department>();
		BindingList<Occupation> occupations = new BindingList<Occupation>();


		public StatTimeTracking()
		{
			InitializeComponent();

			FillDepts();
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


		private void FillDepts()
		{
			try
			{
				string sql = "select * from \"Unit\";";
				if (Connection.get_connect() == null)
					throw new NullReferenceException("Не удалось подключиться к базе данных");


				departments.Add(new Department(-1, ""));

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




		private void StatTimeTracking_FormClosed(object sender, FormClosedEventArgs e)
		{
			Form ifrm = Application.OpenForms[Application.OpenForms.Count - 1];
			ifrm.Show();
		}

		private void btn_buildGraphic_Click(object sender, EventArgs e)
		{

			HashSet<string> marks = new HashSet<string>();
			Dictionary<string, Series> mark2series = new Dictionary<string, Series>();
			Dictionary<string, long> mark2val = new Dictionary<string, long>();

			if (checkBox_Я.Checked)
			{
				marks.Add("Я");
				mark2val.Add("Я", 0);
			}
			if (checkBoxБ.Checked)
			{
				marks.Add("Б");
				mark2val.Add("Б", 0);
			}
			if (checkBoxВ.Checked)
			{
				marks.Add("В");
				mark2val.Add("В", 0);
			}
			if (checkBoxНН.Checked)
			{
				marks.Add("НН");
				mark2val.Add("НН", 0);
			}
			if (checkBoxПР.Checked)
			{
				marks.Add("ПР");
				mark2val.Add("ПР", 0);
			}
			if (checkBoxН.Checked)
			{
				marks.Add("Н");
				mark2val.Add("Н", 0);
			}
			if (checkBoxРВ.Checked)
			{
				marks.Add("РВ");
				mark2val.Add("РВ", 0);
			}
			if (checkBoxС.Checked)
			{
				marks.Add("С");
				mark2val.Add("С", 0);
			}
			if (checkBoxК.Checked)
			{
				marks.Add("К");
				mark2val.Add("К", 0);
			}
			if (checkBoxОТ.Checked)
			{
				marks.Add("ОТ");
				mark2val.Add("ОТ", 0);
			}
			if (checkBoxОД.Checked)
			{
				marks.Add("ОД");
				mark2val.Add("ОД", 0);
			}


			//Настройка графика
			chart1.Series.Clear();



			foreach (string m in marks)
			{
				Series s1 = new Series();
				s1.ChartType = SeriesChartType.Line;
				s1.XValueType = ChartValueType.DateTime;
				s1.Name = m;
				s1.ChartType = SeriesChartType.Spline;
				s1.BorderWidth = 3;

				mark2series.Add(m, s1);
				chart1.Series.Add(s1);
			}
						
			
			chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd";
			chart1.ChartAreas[0].AxisX.Interval = 1;
			chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Days;
			chart1.ChartAreas[0].AxisX.IntervalOffset = 1;
			

			DateTime dateStart = new DateTime((int)yearFrom.Value, (int)monthFrom.Value, 1);
			DateTime dateFinish = new DateTime((int)yearTo.Value, (int)monthTo.Value, DateTime.DaysInMonth((int)yearTo.Value, (int)monthTo.Value));

			chart1.ChartAreas[0].AxisX.Minimum = dateStart.ToOADate();
			chart1.ChartAreas[0].AxisX.Maximum = dateFinish.ToOADate();


			//ПРойти по всем дням
			TimeSpan oneDay = new TimeSpan(1, 0, 0, 0);
			for (DateTime today = dateStart; today < dateFinish; today += oneDay)
			{
				int idUnit = (department.SelectedItem as Department).Id;
				DateTime dateSostav = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month));

				string sql = "select m.\"ShortName\", count(f.\"pk_fact\")" +
						" from \"TimeTracking\" t, \"StringTimeTracking\" s, \"Fact\" f, \"MarkTimeTracking\" m" +
						" where t.\"pk_time_tracking\" = s.\"pk_time_tracking\" " +
						" and s.\"pk_string_time_tracking\" = f.\"pk_string_time_tracking\"" +
						" and f.\"pk_mark_time_tracking\" = m.\"pk_mark_time_tracking\" " +
						" and t.\"date_sostav\" = '" + date2str(dateSostav) + "' " +
						" and f.\"data\" = '" + date2str(today) + "' ";

				if (idUnit != -1)
					sql += " and t.\"pk_unit\" = '" + idUnit + "' " +
						" group by m.\"ShortName\"";
				else
					sql += " group by m.\"ShortName\"";

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

						string curM = (string)obj[0];
						if (marks.Contains(curM)) {
							long newVal = (long)obj[1];
							mark2val[curM] += newVal;
							
							mark2series[curM].Points.AddXY(today, newVal);						
						}
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
				
			}
		}
		
	}
}
