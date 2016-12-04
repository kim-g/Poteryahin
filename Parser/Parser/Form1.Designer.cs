namespace Parser
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
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.PB = new System.Windows.Forms.ProgressBar();
            this.PBL = new System.Windows.Forms.Label();
            this.WorkingTimer = new System.Windows.Forms.Timer(this.components);
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(447, 88);
            this.button1.TabIndex = 2;
            this.button1.Text = "Экспортировать";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(-2, -3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(153, 25);
            this.button2.TabIndex = 3;
            this.button2.Text = "Пересохранить настройки";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.Location = new System.Drawing.Point(16, 268);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(447, 58);
            this.button3.TabIndex = 4;
            this.button3.Text = "Закрыть";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // PB
            // 
            this.PB.Location = new System.Drawing.Point(12, 139);
            this.PB.Maximum = 1000;
            this.PB.Name = "PB";
            this.PB.Size = new System.Drawing.Size(447, 23);
            this.PB.Step = 1;
            this.PB.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.PB.TabIndex = 5;
            this.PB.Visible = false;
            // 
            // PBL
            // 
            this.PBL.AutoSize = true;
            this.PBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.PBL.Location = new System.Drawing.Point(12, 116);
            this.PBL.Name = "PBL";
            this.PBL.Size = new System.Drawing.Size(51, 20);
            this.PBL.TabIndex = 6;
            this.PBL.Text = "label1";
            this.PBL.Visible = false;
            // 
            // WorkingTimer
            // 
            this.WorkingTimer.Enabled = true;
            this.WorkingTimer.Interval = 500;
            this.WorkingTimer.Tick += new System.EventHandler(this.WorkingTimer_Tick);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button4.Location = new System.Drawing.Point(12, 168);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(447, 36);
            this.button4.TabIndex = 7;
            this.button4.Text = "Прервать работу";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Visible = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button5.Location = new System.Drawing.Point(12, 226);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(447, 36);
            this.button5.TabIndex = 8;
            this.button5.Text = "О программе";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 338);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.PBL);
            this.Controls.Add(this.PB);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.DoubleBuffered = true;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Экспорт карточек Билайн из Excel в XML";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ProgressBar PB;
        private System.Windows.Forms.Label PBL;
        private System.Windows.Forms.Timer WorkingTimer;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
    }
}

