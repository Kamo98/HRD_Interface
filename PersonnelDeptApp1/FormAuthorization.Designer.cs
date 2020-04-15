namespace PersonnelDeptApp1
{
	partial class FormAuthorization
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.button1 = new System.Windows.Forms.Button();
			this.tb_login = new System.Windows.Forms.TextBox();
			this.tb_pass = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(148, 154);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(94, 25);
			this.button1.TabIndex = 0;
			this.button1.Text = "Войти";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// tb_login
			// 
			this.tb_login.Location = new System.Drawing.Point(113, 46);
			this.tb_login.Name = "tb_login";
			this.tb_login.Size = new System.Drawing.Size(176, 20);
			this.tb_login.TabIndex = 1;
			this.tb_login.Text = "admin1";
			// 
			// tb_pass
			// 
			this.tb_pass.Location = new System.Drawing.Point(113, 102);
			this.tb_pass.Name = "tb_pass";
			this.tb_pass.PasswordChar = '*';
			this.tb_pass.Size = new System.Drawing.Size(176, 20);
			this.tb_pass.TabIndex = 2;
			this.tb_pass.Text = "admin1";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(53, 49);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 13);
			this.label1.TabIndex = 3;
			this.label1.Text = "Логин:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(53, 105);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Пароль:";
			// 
			// FormAuthorization
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(358, 230);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.tb_pass);
			this.Controls.Add(this.tb_login);
			this.Controls.Add(this.button1);
			this.Name = "FormAuthorization";
			this.Text = "Авторизация";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.TextBox tb_login;
		private System.Windows.Forms.TextBox tb_pass;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
	}
}