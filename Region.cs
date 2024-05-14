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
        public double N2009;// значение численности по годам
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
        public List<Region>regions;
        public void LoadFromExcel()//загрузка данных из Excel
        {
            int rCnt; //количество строк
            regions = new List<Region>();//
            regions.Clear();//очищение

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
                double izm = Math.Round( ((Zn2023 - Zn2009) / Zn2009) * 100,2);

                //создание объекта регион сколько строк столько и регионов
                Region region = new Region();
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

                regions.Add(region);
            }
            //
            ExcelWorkBook.Close(true, null, null); //закрытие книгу
            ExcelApp.Quit();//закрытие excel
        }
    }

}
