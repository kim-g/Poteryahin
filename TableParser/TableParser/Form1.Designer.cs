namespace TableParser
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.PBA = new System.Windows.Forms.ProgressBar();
            this.PBAL = new System.Windows.Forms.Label();
            this.PBC = new System.Windows.Forms.ProgressBar();
            this.PBCL = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(13, 14);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(400, 56);
            this.button1.TabIndex = 0;
            this.button1.Text = "Сделать мне хорошо";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.PBA);
            this.panel1.Controls.Add(this.PBAL);
            this.panel1.Controls.Add(this.PBC);
            this.panel1.Controls.Add(this.PBCL);
            this.panel1.Location = new System.Drawing.Point(12, 78);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(401, 165);
            this.panel1.TabIndex = 1;
            this.panel1.Visible = false;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(111, 128);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(150, 34);
            this.button2.TabIndex = 4;
            this.button2.Text = "Отмена";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // PBA
            // 
            this.PBA.Location = new System.Drawing.Point(3, 85);
            this.PBA.Maximum = 1000;
            this.PBA.Name = "PBA";
            this.PBA.Size = new System.Drawing.Size(395, 23);
            this.PBA.TabIndex = 3;
            // 
            // PBAL
            // 
            this.PBAL.AutoSize = true;
            this.PBAL.Location = new System.Drawing.Point(0, 62);
            this.PBAL.Name = "PBAL";
            this.PBAL.Size = new System.Drawing.Size(53, 20);
            this.PBAL.TabIndex = 2;
            this.PBAL.Text = "Всего";
            // 
            // PBC
            // 
            this.PBC.Location = new System.Drawing.Point(3, 23);
            this.PBC.Maximum = 1000;
            this.PBC.Name = "PBC";
            this.PBC.Size = new System.Drawing.Size(395, 23);
            this.PBC.TabIndex = 1;
            // 
            // PBCL
            // 
            this.PBCL.AutoSize = true;
            this.PBCL.Location = new System.Drawing.Point(0, 0);
            this.PBCL.Name = "PBCL";
            this.PBCL.Size = new System.Drawing.Size(51, 20);
            this.PBCL.TabIndex = 0;
            this.PBCL.Text = "label1";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 500;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 260);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(139, 34);
            this.button3.TabIndex = 2;
            this.button3.Text = "О программе";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(271, 260);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(139, 34);
            this.button4.TabIndex = 3;
            this.button4.Text = "Выход";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(426, 309);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.button1);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Разедлитель книг Excel";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ProgressBar PBA;
        private System.Windows.Forms.Label PBAL;
        private System.Windows.Forms.ProgressBar PBC;
        private System.Windows.Forms.Label PBCL;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
    }
}

