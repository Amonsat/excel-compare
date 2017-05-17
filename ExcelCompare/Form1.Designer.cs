namespace ExcelCompare {
    partial class Form1 {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent() {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbFile1Path = new System.Windows.Forms.TextBox();
            this.bFile1Open = new System.Windows.Forms.Button();
            this.cbFile1SheetList = new System.Windows.Forms.ComboBox();
            this.cbFile1ColumnList = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbFile2ColumnList = new System.Windows.Forms.ComboBox();
            this.cbFile2SheetList = new System.Windows.Forms.ComboBox();
            this.bFile2Open = new System.Windows.Forms.Button();
            this.tbFile2Path = new System.Windows.Forms.TextBox();
            this.bCompare = new System.Windows.Forms.Button();
            this.pbProgress = new System.Windows.Forms.ProgressBar();
            this.lStep = new System.Windows.Forms.Label();
            this.lTime = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cbFile1ColumnList);
            this.groupBox1.Controls.Add(this.cbFile1SheetList);
            this.groupBox1.Controls.Add(this.bFile1Open);
            this.groupBox1.Controls.Add(this.tbFile1Path);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(420, 78);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Первый фаил";
            // 
            // tbFile1Path
            // 
            this.tbFile1Path.Location = new System.Drawing.Point(12, 19);
            this.tbFile1Path.Name = "tbFile1Path";
            this.tbFile1Path.Size = new System.Drawing.Size(315, 20);
            this.tbFile1Path.TabIndex = 0;
            // 
            // bFile1Open
            // 
            this.bFile1Open.Location = new System.Drawing.Point(333, 16);
            this.bFile1Open.Name = "bFile1Open";
            this.bFile1Open.Size = new System.Drawing.Size(75, 23);
            this.bFile1Open.TabIndex = 1;
            this.bFile1Open.Text = "Открыть";
            this.bFile1Open.UseVisualStyleBackColor = true;
            this.bFile1Open.Click += new System.EventHandler(this.bFile1Open_Click);
            // 
            // cbFile1SheetList
            // 
            this.cbFile1SheetList.FormattingEnabled = true;
            this.cbFile1SheetList.Location = new System.Drawing.Point(53, 45);
            this.cbFile1SheetList.Name = "cbFile1SheetList";
            this.cbFile1SheetList.Size = new System.Drawing.Size(121, 21);
            this.cbFile1SheetList.TabIndex = 2;
            this.cbFile1SheetList.SelectedIndexChanged += new System.EventHandler(this.cbFile1SheetList_SelectedIndexChanged);
            // 
            // cbFile1ColumnList
            // 
            this.cbFile1ColumnList.FormattingEnabled = true;
            this.cbFile1ColumnList.Location = new System.Drawing.Point(287, 45);
            this.cbFile1ColumnList.Name = "cbFile1ColumnList";
            this.cbFile1ColumnList.Size = new System.Drawing.Size(121, 21);
            this.cbFile1ColumnList.TabIndex = 3;
            this.cbFile1ColumnList.SelectedIndexChanged += new System.EventHandler(this.cbFile1ColumnList_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Лист";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(180, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Ключевой столбец";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.cbFile2ColumnList);
            this.groupBox2.Controls.Add(this.cbFile2SheetList);
            this.groupBox2.Controls.Add(this.bFile2Open);
            this.groupBox2.Controls.Add(this.tbFile2Path);
            this.groupBox2.Location = new System.Drawing.Point(12, 96);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(420, 78);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Второй фаил";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(180, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(101, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Ключевой столбец";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Лист";
            // 
            // cbFile2ColumnList
            // 
            this.cbFile2ColumnList.FormattingEnabled = true;
            this.cbFile2ColumnList.Location = new System.Drawing.Point(287, 45);
            this.cbFile2ColumnList.Name = "cbFile2ColumnList";
            this.cbFile2ColumnList.Size = new System.Drawing.Size(121, 21);
            this.cbFile2ColumnList.TabIndex = 3;
            this.cbFile2ColumnList.SelectedIndexChanged += new System.EventHandler(this.cbFile2ColumnList_SelectedIndexChanged);
            // 
            // cbFile2SheetList
            // 
            this.cbFile2SheetList.FormattingEnabled = true;
            this.cbFile2SheetList.Location = new System.Drawing.Point(53, 45);
            this.cbFile2SheetList.Name = "cbFile2SheetList";
            this.cbFile2SheetList.Size = new System.Drawing.Size(121, 21);
            this.cbFile2SheetList.TabIndex = 2;
            this.cbFile2SheetList.SelectedIndexChanged += new System.EventHandler(this.cbFile2SheetList_SelectedIndexChanged);
            // 
            // bFile2Open
            // 
            this.bFile2Open.Location = new System.Drawing.Point(333, 16);
            this.bFile2Open.Name = "bFile2Open";
            this.bFile2Open.Size = new System.Drawing.Size(75, 23);
            this.bFile2Open.TabIndex = 1;
            this.bFile2Open.Text = "Открыть";
            this.bFile2Open.UseVisualStyleBackColor = true;
            this.bFile2Open.Click += new System.EventHandler(this.bFile2Open_Click);
            // 
            // tbFile2Path
            // 
            this.tbFile2Path.Location = new System.Drawing.Point(12, 19);
            this.tbFile2Path.Name = "tbFile2Path";
            this.tbFile2Path.Size = new System.Drawing.Size(315, 20);
            this.tbFile2Path.TabIndex = 0;
            // 
            // bCompare
            // 
            this.bCompare.Location = new System.Drawing.Point(177, 180);
            this.bCompare.Name = "bCompare";
            this.bCompare.Size = new System.Drawing.Size(75, 23);
            this.bCompare.TabIndex = 7;
            this.bCompare.Text = "Сравнить";
            this.bCompare.UseVisualStyleBackColor = true;
            this.bCompare.Click += new System.EventHandler(this.bCompare_Click);
            // 
            // pbProgress
            // 
            this.pbProgress.Location = new System.Drawing.Point(12, 208);
            this.pbProgress.Name = "pbProgress";
            this.pbProgress.Size = new System.Drawing.Size(420, 23);
            this.pbProgress.TabIndex = 8;
            // 
            // lStep
            // 
            this.lStep.AutoSize = true;
            this.lStep.Location = new System.Drawing.Point(9, 236);
            this.lStep.Name = "lStep";
            this.lStep.Size = new System.Drawing.Size(27, 13);
            this.lStep.TabIndex = 9;
            this.lStep.Text = "Шаг";
            // 
            // lTime
            // 
            this.lTime.AutoSize = true;
            this.lTime.Location = new System.Drawing.Point(397, 236);
            this.lTime.Name = "lTime";
            this.lTime.Size = new System.Drawing.Size(34, 13);
            this.lTime.TabIndex = 10;
            this.lTime.Text = "00:00";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(445, 258);
            this.Controls.Add(this.lTime);
            this.Controls.Add(this.lStep);
            this.Controls.Add(this.pbProgress);
            this.Controls.Add(this.bCompare);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Excel Compare - Сравнение Excel файлов";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbFile1ColumnList;
        private System.Windows.Forms.ComboBox cbFile1SheetList;
        private System.Windows.Forms.Button bFile1Open;
        private System.Windows.Forms.TextBox tbFile1Path;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbFile2ColumnList;
        private System.Windows.Forms.ComboBox cbFile2SheetList;
        private System.Windows.Forms.Button bFile2Open;
        private System.Windows.Forms.TextBox tbFile2Path;
        private System.Windows.Forms.Button bCompare;
        private System.Windows.Forms.ProgressBar pbProgress;
        private System.Windows.Forms.Label lStep;
        private System.Windows.Forms.Label lTime;
    }
}

