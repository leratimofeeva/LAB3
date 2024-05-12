namespace lab_3
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.RegionName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2009 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2010 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2011 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2012 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2013 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2014 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2015 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2016 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2017 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2018 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2019 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2020 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2021 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2022 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.N2023 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Pizm = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button2 = new System.Windows.Forms.Button();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.button3 = new System.Windows.Forms.Button();
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1262, 601);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1254, 575);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Загрузка данных";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1115, 57);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(121, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Загрузить данные";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.RegionName,
            this.N2009,
            this.N2010,
            this.N2011,
            this.N2012,
            this.N2013,
            this.N2014,
            this.N2015,
            this.N2016,
            this.N2017,
            this.N2018,
            this.N2019,
            this.N2020,
            this.N2021,
            this.N2022,
            this.N2023,
            this.Pizm});
            this.dataGridView1.Location = new System.Drawing.Point(6, 6);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1083, 539);
            this.dataGridView1.TabIndex = 0;
            //this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // RegionName
            // 
            this.RegionName.HeaderText = "RegionName";
            this.RegionName.Name = "RegionName";
            // 
            // N2009
            // 
            this.N2009.HeaderText = "2009";
            this.N2009.Name = "N2009";
            // 
            // N2010
            // 
            this.N2010.HeaderText = "2010";
            this.N2010.Name = "N2010";
            // 
            // N2011
            // 
            this.N2011.HeaderText = "2011";
            this.N2011.Name = "N2011";
            // 
            // N2012
            // 
            this.N2012.HeaderText = "2012";
            this.N2012.Name = "N2012";
            // 
            // N2013
            // 
            this.N2013.HeaderText = "2013";
            this.N2013.Name = "N2013";
            // 
            // N2014
            // 
            this.N2014.HeaderText = "2014";
            this.N2014.Name = "N2014";
            // 
            // N2015
            // 
            this.N2015.HeaderText = "2015";
            this.N2015.Name = "N2015";
            // 
            // N2016
            // 
            this.N2016.HeaderText = "2016";
            this.N2016.Name = "N2016";
            // 
            // N2017
            // 
            this.N2017.HeaderText = "2017";
            this.N2017.Name = "N2017";
            // 
            // N2018
            // 
            this.N2018.HeaderText = "2018";
            this.N2018.Name = "N2018";
            // 
            // N2019
            // 
            this.N2019.HeaderText = "2019";
            this.N2019.Name = "N2019";
            // 
            // N2020
            // 
            this.N2020.HeaderText = "2020";
            this.N2020.Name = "N2020";
            // 
            // N2021
            // 
            this.N2021.HeaderText = "2021";
            this.N2021.Name = "N2021";
            // 
            // N2022
            // 
            this.N2022.HeaderText = "2022";
            this.N2022.Name = "N2022";
            // 
            // N2023
            // 
            this.N2023.HeaderText = "2023";
            this.N2023.Name = "N2023";
            // 
            // Pizm
            // 
            this.Pizm.HeaderText = "Pizm";
            this.Pizm.Name = "Pizm";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.button2);
            this.tabPage2.Controls.Add(this.chart1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1254, 575);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Графики";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(1126, 103);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(122, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Построить графики";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(19, 22);
            this.chart1.Name = "chart1";
            this.chart1.Size = new System.Drawing.Size(1088, 547);
            this.chart1.TabIndex = 0;
            this.chart1.Text = "chart1";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.numericUpDown2);
            this.tabPage3.Controls.Add(this.label2);
            this.tabPage3.Controls.Add(this.label1);
            this.tabPage3.Controls.Add(this.numericUpDown1);
            this.tabPage3.Controls.Add(this.button3);
            this.tabPage3.Controls.Add(this.chart2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1254, 575);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Экстраполяция";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // numericUpDown2
            // 
            this.numericUpDown2.Location = new System.Drawing.Point(614, 203);
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(120, 20);
            this.numericUpDown2.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(614, 171);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "число n (период)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(614, 114);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Ввести количество N лет";
            //this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(614, 133);
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(120, 20);
            this.numericUpDown1.TabIndex = 2;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(614, 72);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(174, 23);
            this.button3.TabIndex = 1;
            this.button3.Text = "Найти и экстраполировать";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // chart2
            // 
            chartArea2.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea2);
            legend2.Name = "Legend1";
            this.chart2.Legends.Add(legend2);
            this.chart2.Location = new System.Drawing.Point(18, 21);
            this.chart2.Name = "chart2";
            this.chart2.Size = new System.Drawing.Size(431, 346);
            this.chart2.TabIndex = 0;
            this.chart2.Text = "chart2";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1277, 616);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewTextBoxColumn RegionName;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2009;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2010;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2011;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2012;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2013;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2014;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2015;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2016;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2017;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2018;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2019;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2020;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2021;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2022;
        private System.Windows.Forms.DataGridViewTextBoxColumn N2023;
        private System.Windows.Forms.DataGridViewTextBoxColumn Pizm;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private System.Windows.Forms.Label label2;
    }
}

