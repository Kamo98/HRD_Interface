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



		private void Form2_Load(object sender, EventArgs e)
        {
            //AutoCompleteStringCollection listUnit_1 = new AutoCompleteStringCollection();
            //NpgsqlCommand com = new NpgsqlCommand("SELECT \"Name\" FROM \"Unit\"", npgSqlConnection);
            //NpgsqlDataReader reader = com.ExecuteReader();

            //if (reader.HasRows)
            //{
            //    foreach (DbDataRecord rec in reader)
            //    {
            //        listUnit_1.Add(rec.GetString(0));
            //    }

            //}
            //reader.Close();


			FillDepts();	


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
            //AutoCompleteStringCollection listPos = new AutoCompleteStringCollection();
            //try
            //{
            //    string sql = "select \"pos\".\"pk_position\", \"pos\".\"Name\", \"pos\".\"Rate\""
            //                    + " from \"Position\" as \"pos\", \"Unit\" as \"un\""
            //                    + " where \"pos\".\"pk_unit\" = \"un\".\"pk_unit\" and \"un\".\"pk_unit\" = " + ((sender as ComboBox).SelectedItem as Department).Id + ";";
            //    NpgsqlCommand command = new NpgsqlCommand(sql, Connection.get_connect());
            //    NpgsqlDataReader reader = command.ExecuteReader();
            //    foreach (DbDataRecord record in reader)
            //    {
            //        listPos.Add(record.GetString(0));
            //    }
            //    reader.Close();

            //    comboBox2.DataSource = listPos;
            //    comboBox2.AutoCompleteMode = AutoCompleteMode.Suggest;
            //    comboBox2.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //    comboBox2.AutoCompleteCustomSource = listPos;

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}


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

				comboBox2.DataSource = occupations;
				comboBox2.DisplayMember = "Name";

			}
			catch (Exception ex)
			{
				MessageBox.Show("Неизвестная ошибка.\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

		}
	}
}
