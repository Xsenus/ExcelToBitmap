namespace ExcelToBitmap
{
    partial class ExcelToBitmap
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblExcel = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.txtExcel = new System.Windows.Forms.TextBox();
            this.txtPathOut = new System.Windows.Forms.TextBox();
            this.btnAddPathOut = new System.Windows.Forms.Button();
            this.lblPathOut = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblExcel
            // 
            this.lblExcel.AutoSize = true;
            this.lblExcel.Location = new System.Drawing.Point(64, 15);
            this.lblExcel.Name = "lblExcel";
            this.lblExcel.Size = new System.Drawing.Size(68, 13);
            this.lblExcel.TabIndex = 0;
            this.lblExcel.Text = "Файл Excel:";
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(494, 12);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(20, 20);
            this.btnExcel.TabIndex = 1;
            this.btnExcel.Text = "+";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // txtExcel
            // 
            this.txtExcel.Location = new System.Drawing.Point(138, 12);
            this.txtExcel.Name = "txtExcel";
            this.txtExcel.ReadOnly = true;
            this.txtExcel.Size = new System.Drawing.Size(350, 20);
            this.txtExcel.TabIndex = 2;
            // 
            // txtPathOut
            // 
            this.txtPathOut.Location = new System.Drawing.Point(138, 38);
            this.txtPathOut.Name = "txtPathOut";
            this.txtPathOut.ReadOnly = true;
            this.txtPathOut.Size = new System.Drawing.Size(350, 20);
            this.txtPathOut.TabIndex = 5;
            // 
            // btnAddPathOut
            // 
            this.btnAddPathOut.Location = new System.Drawing.Point(494, 38);
            this.btnAddPathOut.Name = "btnAddPathOut";
            this.btnAddPathOut.Size = new System.Drawing.Size(20, 20);
            this.btnAddPathOut.TabIndex = 4;
            this.btnAddPathOut.Text = "+";
            this.btnAddPathOut.UseVisualStyleBackColor = true;
            this.btnAddPathOut.Click += new System.EventHandler(this.btnAddPathOut_Click);
            // 
            // lblPathOut
            // 
            this.lblPathOut.AutoSize = true;
            this.lblPathOut.Location = new System.Drawing.Point(12, 41);
            this.lblPathOut.Name = "lblPathOut";
            this.lblPathOut.Size = new System.Drawing.Size(120, 13);
            this.lblPathOut.TabIndex = 3;
            this.lblPathOut.Text = "Конечная директория:";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(433, 64);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(81, 29);
            this.btnStart.TabIndex = 6;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // ExcelToBitmap
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(534, 99);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.txtPathOut);
            this.Controls.Add(this.btnAddPathOut);
            this.Controls.Add(this.lblPathOut);
            this.Controls.Add(this.txtExcel);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.lblExcel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "ExcelToBitmap";
            this.Text = "Excel To Bitmap (ilel@list.ru)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblExcel;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.TextBox txtExcel;
        private System.Windows.Forms.TextBox txtPathOut;
        private System.Windows.Forms.Button btnAddPathOut;
        private System.Windows.Forms.Label lblPathOut;
        private System.Windows.Forms.Button btnStart;
    }
}

