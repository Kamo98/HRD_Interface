namespace PersonnelDeptApp1
{
    partial class FormOrderSearch
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.empListCB = new System.Windows.Forms.ComboBox();
            this.orderTypesCB = new System.Windows.Forms.ComboBox();
            this.selectedEmpId = new System.Windows.Forms.Label();
            this.ordersTable = new System.Windows.Forms.DataGridView();
            this.orderIdCol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OrderNumCol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OrderDateCol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PositionCol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ContractNum = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ShowOrderBtnCol = new System.Windows.Forms.DataGridViewButtonColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.FindBTN = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ordersTable)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayoutPanel1.Controls.Add(this.empListCB, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.orderTypesCB, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.selectedEmpId, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.ordersTable, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.FindBTN, 2, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 85F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(865, 650);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // empListCB
            // 
            this.empListCB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.empListCB.Font = new System.Drawing.Font("Times New Roman", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.empListCB.FormattingEnabled = true;
            this.empListCB.Location = new System.Drawing.Point(3, 35);
            this.empListCB.Name = "empListCB";
            this.empListCB.Size = new System.Drawing.Size(513, 27);
            this.empListCB.TabIndex = 0;
            this.empListCB.SelectedIndexChanged += new System.EventHandler(this.empListCB_SelectedIndexChanged);
            // 
            // orderTypesCB
            // 
            this.orderTypesCB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.orderTypesCB.Font = new System.Drawing.Font("Times New Roman", 10.2F);
            this.orderTypesCB.FormattingEnabled = true;
            this.orderTypesCB.Location = new System.Drawing.Point(608, 35);
            this.orderTypesCB.Name = "orderTypesCB";
            this.orderTypesCB.Size = new System.Drawing.Size(254, 27);
            this.orderTypesCB.TabIndex = 1;
            // 
            // selectedEmpId
            // 
            this.selectedEmpId.AutoSize = true;
            this.selectedEmpId.Dock = System.Windows.Forms.DockStyle.Fill;
            this.selectedEmpId.Font = new System.Drawing.Font("Times New Roman", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.selectedEmpId.Location = new System.Drawing.Point(522, 32);
            this.selectedEmpId.Name = "selectedEmpId";
            this.selectedEmpId.Size = new System.Drawing.Size(80, 32);
            this.selectedEmpId.TabIndex = 2;
            this.selectedEmpId.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ordersTable
            // 
            this.ordersTable.AllowUserToAddRows = false;
            this.ordersTable.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.ordersTable.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.ordersTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ordersTable.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.orderIdCol,
            this.OrderNumCol,
            this.OrderDateCol,
            this.PositionCol,
            this.ContractNum,
            this.ShowOrderBtnCol});
            this.tableLayoutPanel1.SetColumnSpan(this.ordersTable, 3);
            this.ordersTable.Location = new System.Drawing.Point(3, 99);
            this.ordersTable.Name = "ordersTable";
            this.ordersTable.ReadOnly = true;
            this.ordersTable.RowHeadersVisible = false;
            this.ordersTable.RowHeadersWidth = 51;
            this.ordersTable.RowTemplate.Height = 24;
            this.ordersTable.Size = new System.Drawing.Size(859, 539);
            this.ordersTable.TabIndex = 3;
            this.ordersTable.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ordersTable_CellContentClick);
            // 
            // orderIdCol
            // 
            this.orderIdCol.HeaderText = "ID";
            this.orderIdCol.MinimumWidth = 6;
            this.orderIdCol.Name = "orderIdCol";
            this.orderIdCol.ReadOnly = true;
            this.orderIdCol.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.orderIdCol.Visible = false;
            this.orderIdCol.Width = 125;
            // 
            // OrderNumCol
            // 
            this.OrderNumCol.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.OrderNumCol.DefaultCellStyle = dataGridViewCellStyle2;
            this.OrderNumCol.FillWeight = 15F;
            this.OrderNumCol.HeaderText = "Номер приказа";
            this.OrderNumCol.MinimumWidth = 6;
            this.OrderNumCol.Name = "OrderNumCol";
            this.OrderNumCol.ReadOnly = true;
            // 
            // OrderDateCol
            // 
            this.OrderDateCol.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.OrderDateCol.DefaultCellStyle = dataGridViewCellStyle3;
            this.OrderDateCol.FillWeight = 15F;
            this.OrderDateCol.HeaderText = "Дата приказа";
            this.OrderDateCol.MinimumWidth = 6;
            this.OrderDateCol.Name = "OrderDateCol";
            this.OrderDateCol.ReadOnly = true;
            // 
            // PositionCol
            // 
            this.PositionCol.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.PositionCol.DefaultCellStyle = dataGridViewCellStyle4;
            this.PositionCol.FillWeight = 40F;
            this.PositionCol.HeaderText = "Должность";
            this.PositionCol.MinimumWidth = 6;
            this.PositionCol.Name = "PositionCol";
            this.PositionCol.ReadOnly = true;
            // 
            // ContractNum
            // 
            this.ContractNum.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.ContractNum.DefaultCellStyle = dataGridViewCellStyle5;
            this.ContractNum.FillWeight = 15F;
            this.ContractNum.HeaderText = "Номер договора";
            this.ContractNum.MinimumWidth = 6;
            this.ContractNum.Name = "ContractNum";
            this.ContractNum.ReadOnly = true;
            // 
            // ShowOrderBtnCol
            // 
            this.ShowOrderBtnCol.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ShowOrderBtnCol.FillWeight = 15F;
            this.ShowOrderBtnCol.HeaderText = "";
            this.ShowOrderBtnCol.MinimumWidth = 6;
            this.ShowOrderBtnCol.Name = "ShowOrderBtnCol";
            this.ShowOrderBtnCol.ReadOnly = true;
            this.ShowOrderBtnCol.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(3, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(513, 19);
            this.label1.TabIndex = 4;
            this.label1.Text = "Сотрудник";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(608, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(254, 19);
            this.label2.TabIndex = 5;
            this.label2.Text = "Тип приказа";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // FindBTN
            // 
            this.FindBTN.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FindBTN.Location = new System.Drawing.Point(608, 67);
            this.FindBTN.Name = "FindBTN";
            this.FindBTN.Size = new System.Drawing.Size(254, 26);
            this.FindBTN.TabIndex = 6;
            this.FindBTN.Text = "Найти приказы";
            this.FindBTN.UseVisualStyleBackColor = true;
            this.FindBTN.Click += new System.EventHandler(this.FindBTN_Click);
            // 
            // FormOrderSearch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(865, 650);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormOrderSearch";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Поиск приказов";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormOrderSearch_FormClosed);
            this.Load += new System.EventHandler(this.FormOrderSearch_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ordersTable)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ComboBox empListCB;
        private System.Windows.Forms.ComboBox orderTypesCB;
        private System.Windows.Forms.Label selectedEmpId;
        private System.Windows.Forms.DataGridView ordersTable;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button FindBTN;
        private System.Windows.Forms.DataGridViewTextBoxColumn orderIdCol;
        private System.Windows.Forms.DataGridViewTextBoxColumn OrderNumCol;
        private System.Windows.Forms.DataGridViewTextBoxColumn OrderDateCol;
        private System.Windows.Forms.DataGridViewTextBoxColumn PositionCol;
        private System.Windows.Forms.DataGridViewTextBoxColumn ContractNum;
        private System.Windows.Forms.DataGridViewButtonColumn ShowOrderBtnCol;
    }
}