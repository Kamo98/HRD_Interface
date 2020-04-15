using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PersonnelDeptApp1
{
	public partial class FormAuthorization : Form
	{
		public FormAuthorization()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			bool correctConn = Connection.create_instance(tb_login.Text, tb_pass.Text);

			if (correctConn)
			{
				//MessageBox.Show(Connection.get_role_str());
				Form2 mainForm = new Form2(this);
				this.Hide();
				mainForm.Show();
			}

			
		}		
	}
}
