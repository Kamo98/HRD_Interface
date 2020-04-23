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
   

    public partial class Form2 : Form
    {
        NpgsqlConnection npgSqlConnection;

        private FormAuthorization formAuthorization;

		BindingList<Department> departments = new BindingList<Department>();
		BindingList<Occupation> occupations = new BindingList<Occupation>();

		public Form2(FormAuthorization formAuthorization)
        {
			InitializeComponent();
			this.formAuthorization = formAuthorization;

            npgSqlConnection = Connection.get_connect();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void приказыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 f4 = new Form4();
            f4.Show();
			this.Hide();
        }

		private void выйтиToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void Form2_FormClosing(object sender, FormClosingEventArgs e)
		{
			DialogResult dialogResult = MessageBox.Show("Уверены, что хотите выйти из системы?", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);

			if (dialogResult == DialogResult.Yes)
			{
				Connection.close_connection();
				formAuthorization.Show();
			}
			else
				e.Cancel = true;
		}

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void отчётыToolStripMenuItem_Click(object sender, EventArgs e)
        {

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

				departments.Add(new Department(-1, ""));

				foreach (DbDataRecord record in reader)
				{
					Department newDept = new Department((int)record["pk_unit"], (string)record["Name"]);
					departments.Add(newDept);

				}
				reader.Close();

				comboBox1.DataSource = departments;
				comboBox1.DisplayMember = "Name";
				comboBox1.DataSource = departments;
				comboBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
				comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;

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


		private void FillGrid()
		{
			try
			{
				string baseStr = "select c.\"pk_personal_card\", c.\"surname\", c.\"name\", c.\"otchestvo\", u.\"Name\", d.\"Name\" " +
				" from \"PersonalCard\" c, \"PeriodPosition\" p, \"Position\" d, \"Unit\" u" +
				" where c.\"pk_personal_card\" = p.\"pk_personal_card\" " +
				" and d.\"pk_position\" = p.\"pk_position\" " +
				" and u.\"pk_unit\" = d.\"pk_unit\" " +
				" and p.\"DateTo\" is null ";


				if ((comboBox1.SelectedItem as Department).Id != -1)
				{
					int idUnit = (comboBox1.SelectedItem as Department).Id;
					baseStr += " and d.\"pk_unit\" ='" + idUnit + "' ";
				}
				if ((comboBox2.SelectedItem as Occupation).Id != -1)
				{
					int idPosition = (comboBox2.SelectedItem as Occupation).Id;
					baseStr += " and p.\"pk_position\" ='" + idPosition + "' ";
				}
				if (richTextBox1.Text != "")
				{
					string surename = richTextBox1.Text.Trim();
					baseStr += " and c.\"surname\" ='" + surename + "' ";
				}
				if (richTextBox2.Text != "")
				{
					string name = richTextBox2.Text.Trim();
					baseStr += " and c.\"name\" ='" + name + "' ";
				}
				if (richTextBox3.Text != "")
				{
					string otchestvo = richTextBox3.Text.Trim();
					baseStr += " and c.\"otchestvo\" ='" + otchestvo + "' ";
				}


				if (Connection.get_connect() == null)
					throw new NullReferenceException("Не удалось подключиться к базе данных");

				NpgsqlCommand command = new NpgsqlCommand(baseStr, Connection.get_connect());
				NpgsqlDataReader reader = command.ExecuteReader();
				int k = 0;
				foreach (DbDataRecord record in reader)
				{
					object[] obj = new object[record.FieldCount];
					record.GetValues(obj);

					dataGridView1.Rows.Add();
					for (int i = 0; i < 6; i++)
					{
						dataGridView1.Rows[k].Cells[i].Value = obj[i].ToString();
					}
					k++;

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

		private void Form2_Load(object sender, EventArgs e)
        {
           
			FillDepts();
			FillGrid();

		}

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
          
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }


		private void FillOccups(object sender, EventArgs e)
		{
			
		}


		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
			occupations.Clear();

			try
			{
				string sql = "select \"pos\".\"pk_position\", \"pos\".\"Name\", \"pos\".\"Rate\""
								+ " from \"Position\" as \"pos\", \"Unit\" as \"un\""
								+ " where \"pos\".\"pk_unit\" = \"un\".\"pk_unit\" and \"un\".\"pk_unit\" = " + ((sender as ComboBox).SelectedItem as Department).Id + ";";
				NpgsqlCommand command = new NpgsqlCommand(sql, Connection.get_connect());
				NpgsqlDataReader reader = command.ExecuteReader();

				occupations.Add(new Occupation(-1, ""));

				foreach (DbDataRecord record in reader)
				{
					occupations.Add(new Occupation((int)record["pk_position"], (string)record["Name"], (decimal)record["Rate"]));
				}
				reader.Close();

				comboBox2.DataSource = occupations;
				comboBox2.DisplayMember = "Name";

			}
			catch (Exception ex)
			{
				MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

		}

	
		private void button4_Click(object sender, EventArgs e)
		{
			button4.Enabled = false;
			dataGridView1.Rows.Clear(); //очещаем грид

			FillGrid();

			button4.Enabled = true;
		}
            
    }

}
