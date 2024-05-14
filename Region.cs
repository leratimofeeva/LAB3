using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace lab_3
{
    public abstract class Territoriya //абстрактный класс Территория
    {
        public string RegionName;//свойство
    }
    public class Region: Territoriya // класс Регион наследует свойства от абстрактного класса Территория
    {
        public double N2009;// свойства в которых может быть значение численности по годам 
        public double N2010;
        public double N2011;
        public double N2012;
        public double N2013;
        public double N2014;
        public double N2015;
        public double N2016;
        public double N2017;
        public double N2018;
        public double N2019;
        public double N2020;
        public double N2021;
        public double N2022;
        public double N2023;

        public double Pizm; //процент изменения (Конечное значение - Начальное значение) ⁄ Начальное значение × 100
    }
    public class ListRegion
    {
        public List<Region> regions;
        public void LoadFromExcel()//загрузка данных из Excel
        {
            int rCnt; //количество строк
            regions = new List<Region>();//создаем список типа регион
            regions.Clear();//очищение списка

            OpenFileDialog opf = new OpenFileDialog();//открытие окна для выбора файла
            opf.ShowDialog();
            string filename = opf.FileName; // имя открытого файла
            //открытие Excel подключения Excel в ссылках
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelRange;

            ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                false, 0, true, 1, 0);// открытие Excel по пути файла


            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelRange = ExcelWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)// цикл от первой строки по количеству строк
            {
                double Zn2009 = (ExcelRange.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2; //получение значения ячейки

                double Zn2023 = (ExcelRange.Cells[rCnt, 16] as Microsoft.Office.Interop.Excel.Range).Value2;

                // изменение численности за 15 летв процентном соотношении.Последний год минус первый, делим на первый,умнож на100
                double izm = Math.Round(((Zn2023 - Zn2009) / Zn2009) * 100, 2);

                //создание объекта регион сколько строк столько и регионов
                Region region = new Region();
                //заполнение свойства объекта регион
                region.RegionName = (string)(ExcelRange.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2;

                region.N2009 = Zn2009;
                region.N2010 = (ExcelRange.Cells[rCnt, 3] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2011 = (ExcelRange.Cells[rCnt, 4] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2012 = (ExcelRange.Cells[rCnt, 5] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2013 = (ExcelRange.Cells[rCnt, 6] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2014 = (ExcelRange.Cells[rCnt, 7] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2015 = (ExcelRange.Cells[rCnt, 8] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2016 = (ExcelRange.Cells[rCnt, 9] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2017 = (ExcelRange.Cells[rCnt, 10] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2018 = (ExcelRange.Cells[rCnt, 11] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2019 = (ExcelRange.Cells[rCnt, 12] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2020 = (ExcelRange.Cells[rCnt, 13] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2021 = (ExcelRange.Cells[rCnt, 14] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2022 = (ExcelRange.Cells[rCnt, 15] as Microsoft.Office.Interop.Excel.Range).Value2;
                region.N2023 = Zn2023;

                region.Pizm = izm;
                //заполнение свойства объекта регион

                regions.Add(region);//добавление объекта регион в список
            }
            ExcelWorkBook.Close(true, null, null); //закрытие книгу
            ExcelApp.Quit();//закрытие excel
        }

        public List<Region> copy()//создание списка с двумя строками
        {
            List<Region> regionsotr = new List<Region>();//список регионов с отрицательным процентом изменения 
            List<Region> regionsotr2 = new List<Region>();//список регионов с наименьшим и наибольшим отрицательным процентом изменения

            //перебираем строки списка regions созданного раннее чтобы заполнить другие два списка
            foreach (Region region in regions)//для каждой строки из заполненного списка регионов 
            {
                if (region.Pizm < 0)//проверяем, если процент изменения отрицательный,
                {
                    regionsotr.Add(region);//то добавляем строку в новый список где будут храниться строки только с отрицательным значением 
                }
                
            }
            regionsotr.Sort((x, y) => x.Pizm.CompareTo(y.Pizm));//сортировка списка от наименьшего отрицательного значения к наибольшему

            regionsotr2.Add(regionsotr[0]);//добавляем первую строку из списка в которой хранится наименьшее отрицательное
            regionsotr2.Add(regions[regionsotr.Count - 1]);//добавляем последнюю строку из списка regionsotr в которой хранится наибольшее отрицательное 
            return regionsotr2;
        }
        public double extrapol(List<double> massEkstp, int Kol)//экстраполяция
        {
            double y = 0;

            foreach (var LR in massEkstp)
            {
                y = y + LR;
            }

            y = y / Kol;

            return y;
        }
    }

}
