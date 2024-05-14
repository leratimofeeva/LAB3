using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace lab_3
{
    public partial class Form1 : Form
    {
        ListRegion listRegion = new ListRegion(); // создание объекта список регионов
        public Form1()
        {
            InitializeComponent();
        }

        //загрузка данных из Excel
        private void button1_Click(object sender, EventArgs e)
        {
            //ListRegion listRegion = new ListRegion();
            listRegion.LoadFromExel(); // загрузка из экселя

            // очищение dataGridView1
            dataGridView1.Rows.Clear();

            foreach (var LR in listRegion.regions)
            {
                //dataGridView1.Rows.Add(LR); // столбцы называются так же как и свойства
                dataGridView1.Rows.Add(LR.RegionName, LR.N2009, LR.N2010, LR.N2011, LR.N2012, LR.N2013, LR.N2014, LR.N2015, LR.N2016, LR.N2017, LR.N2018, LR.N2019, LR.N2020, LR.N2021, LR.N2022, LR.N2023, LR.Pizm);
            }
        }

        //построить графики по данным таблицы
        private void button2_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            Random rnd = new Random();
            //chart1.Series[2].Points.Clear(); // очищение точек

            //MessageBox.Show("Число точек должно быть не меньше 4");
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //for (int i = 1; i < 86; i++)
            {
                int r = rnd.Next(0, 255);
                int g = rnd.Next(0, 255);
                int b = rnd.Next(0, 255);

                chart1.Series.Add("Series" + i.ToString());
                chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);

                chart1.Series["Series" + i.ToString()].LegendText = dataGridView1.Rows[i].Cells[0].Value.ToString(); //"Название Series"
                                                                                                                     //MessageBox.Show("Назавание " + dataGridView1.Rows[i].Cells[0].Value);



                for (int k = 1; k < 16; k++)
                {
                    int x = 2008 + k;
                    int y = Convert.ToInt32(dataGridView1.Rows[i].Cells[k].Value);
                    chart1.Series["Series" + i.ToString()].Points.AddXY(x, y);
                    // тип графика или в свойствах
                    chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                    //chart1.Series["Series" + i.ToString()].Points.AddXY(x, y);
                }
            }

            // Создание оси X
            chart1.ChartAreas[0].AxisX.Minimum = 2009;
            chart1.ChartAreas[0].AxisX.Maximum = 2023; // Диапазон значений X от 2008 до 2023;
            chart1.ChartAreas[0].AxisX.Interval = 1; // Шаг значения X

            //// Создание оси Y
            chart1.ChartAreas[0].AxisY.Minimum = 0; // Минимальное значение Y
            //chart1.ChartAreas[0].AxisY.Maximum = 12000000; // Максимальное значение Y
            //chart1.ChartAreas[0].AxisY.Interval = 1; // Шаг значения Y
        }

        // экстраполяция
        private void button3_Click(object sender, EventArgs e)
        {
            int NYears_ = Convert.ToInt32(numericUpDown1.Value);  // количество лет для экстраполяции
            int kol = Convert.ToInt32(numericUpDown2.Value);   // интервал времени

            List<Region> copy = new List<Region>();

            copy = listRegion.copyOrder(); // поиск двух нужных строк

            //dataGridView2.Rows.Clear();

            //foreach (var LR in copy)
            //{
            //    //dataGridView1.Rows.Add(LR); // столбцы называются так же как и свойства
            //    dataGridView2.Rows.Add(LR.RegionName, LR.N2009, LR.N2010, LR.N2011, LR.N2012, LR.N2013, LR.N2014,
            //        LR.N2015, LR.N2016, LR.N2017, LR.N2018, LR.N2019, LR.N2020, LR.N2021, LR.N2022, LR.N2023, LR.Pizm);
            //}


            chart2.Series.Clear();
            Random rnd = new Random();

            int i = 0;
            //MessageBox.Show("Число точек должно быть не меньше 4");

            foreach (var LR in copy) // построение графика этих строк
            {
                // массив для экстраполяции

                List<double> massEkstp = new List<double>();
                i++;
                int r = rnd.Next(0, 255);
                int g = rnd.Next(0, 255);
                int b = rnd.Next(0, 255);

                chart2.Series.Add("Series" + i.ToString());
                //chart2.Series.Add("Series" + i.ToString());
                chart2.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                chart2.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);

                chart2.Series["Series" + i.ToString()].LegendText = LR.RegionName; //"Название Series"

                // другие цвета  график экстраполяции
                r = rnd.Next(0, 255);
                g = rnd.Next(0, 255);
                b = rnd.Next(0, 255);
                chart2.Series.Add("Series" + (i + 2).ToString());
                chart2.Series["Series" + (i + 2).ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                chart2.Series["Series" + (i + 2).ToString()].Color = Color.FromArgb(255, r, g, b);
                chart2.Series["Series" + (i + 2).ToString()].LegendText = LR.RegionName + " экстрапол";

                //MessageBox.Show("Назавание " + dataGridView1.Rows[i].Cells[0].Value);
                //massEkstp.Add((2010, LR.N2010));
                // вывод текущих точек на график и заполнение массива для экстраполяции
                chart2.Series["Series" + i.ToString()].Points.AddXY(2009, LR.N2009); massEkstp.Add(LR.N2009);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2010, LR.N2010); massEkstp.Add(LR.N2010);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2011, LR.N2011); massEkstp.Add(LR.N2011);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2012, LR.N2012); massEkstp.Add(LR.N2012);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2013, LR.N2013); massEkstp.Add(LR.N2013);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2014, LR.N2014); massEkstp.Add(LR.N2014);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2015, LR.N2015); massEkstp.Add(LR.N2015);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2016, LR.N2016); massEkstp.Add(LR.N2016);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2017, LR.N2017); massEkstp.Add(LR.N2017);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2018, LR.N2018); massEkstp.Add(LR.N2018);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2019, LR.N2019); massEkstp.Add(LR.N2019);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2020, LR.N2020); massEkstp.Add(LR.N2020);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2021, LR.N2021); massEkstp.Add(LR.N2021);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2022, LR.N2022); massEkstp.Add(LR.N2022);
                chart2.Series["Series" + i.ToString()].Points.AddXY(2023, LR.N2023); massEkstp.Add(LR.N2023);
                // последнюю точку добавим на график экстраполяции
                chart2.Series["Series" + (i + 2).ToString()].Points.AddXY(2023, LR.N2023); massEkstp.Add(LR.N2023);
                // тип графика или в свойствах
                //chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;

                // Экстраполяция
                int kolmass = massEkstp.Count;
                while (kol < kolmass) // удаляем ненужные сначала элементы для экстраполяции на необходимое количество лет
                {
                    massEkstp.RemoveAt(0);// уменьшаем количество элементов в массиве
                    kolmass--; // уменьшаем количество элементов в массиве
                }


                for (int k = 1; k <= NYears_; k++) // цикл по количеству лет
                {
                    int x = 2023 + k; // год
                    double y = Convert.ToInt32(listRegion.extrapol(massEkstp, kol)); // численность,  получение по экстраполяции

                    // удалить первый элемент
                    massEkstp.RemoveAt(0);
                    // добавить новый элемент в массив экстраполяции
                    massEkstp.Add(y);

                    chart2.Series["Series" + (i + 2).ToString()].Points.AddXY(x, y);

                    // тип графика или в свойствах
                    //chart2.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;

                }
            }

            // Создание оси X
            chart2.ChartAreas[0].AxisX.Minimum = 2009;
            chart2.ChartAreas[0].AxisX.Maximum = 2023 + NYears_; // Диапазон значений X от 2008 до 2023 + NYears_;
            chart2.ChartAreas[0].AxisX.Interval = 1; // Шаг значения X

            //// Создание оси Y
            chart2.ChartAreas[0].AxisY.Minimum = 0; // Минимальное значение 
        }
    }
}
