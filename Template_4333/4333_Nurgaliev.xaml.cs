using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4333

{
    /// <summary>
    /// Логика взаимодействия для _4333_Nurgaliev.xaml
    /// </summary>
    public partial class _4333_Nurgaliev : Window
    {
        public _4333_Nurgaliev()
        {
            InitializeComponent();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (laba2Entities4 userEntities = new laba2Entities4())
            {
                for (int i = 1; i < _rows; i++)
                {
                    userEntities.isrpo2.Add(new isrpo2()
                    {
                        IdServices = list[i, 0],
                        NameServices = list[i, 1],
                        TypeOfService = list[i, 2],
                        CodeService = list[i, 3],
                        Cost = list[i, 4],
                    });
                }
                userEntities.SaveChanges();
            }
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<isrpo2> AllService;
            using (laba2Entities4 UserEntities = new laba2Entities4())
            {
                AllService = UserEntities.isrpo2.ToList();
            }
            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;
            Excel.Worksheet worksheet1 = app.Worksheets.Add();
            worksheet1.Name = "Категория 1";
            Excel.Worksheet worksheet2 = app.Worksheets.Add();
            worksheet2.Name = "Категория 2";
            Excel.Worksheet worksheet3 = app.Worksheets.Add();
            worksheet3.Name = "Категория 3";
            worksheet1.Cells[1, 1] = "id";
            worksheet1.Cells[1, 2] = "Nazvanie Uslugi";
            worksheet1.Cells[1, 3] = "Vid Uslugi";
            worksheet1.Cells[1, 4] = "Stoimost";

            worksheet2.Cells[1, 1] = "id";
            worksheet2.Cells[1, 2] = "Nazvanie Uslugi";
            worksheet2.Cells[1, 3] = "Vid Uslugi";
            worksheet2.Cells[1, 4] = "Stoimost";

            worksheet3.Cells[1, 1] = "id";
            worksheet3.Cells[1, 2] = "Nazvanie Uslugi";
            worksheet3.Cells[1, 3] = "Vid Uslugi";
            worksheet3.Cells[1, 4] = "Stoimost";
            int rowindex1 = 2;
            int rowindex2 = 2;
            int rowindex3 = 2;

            foreach (var service in AllService)
            {
                if (Convert.ToDouble(service.Cost) < 350)
                {
                    worksheet1.Cells[rowindex1, 1] = service.IdServices;
                    worksheet1.Cells[rowindex1, 2] = service.NameServices;
                    worksheet1.Cells[rowindex1, 3] = service.TypeOfService;
                    worksheet1.Cells[rowindex1, 4] = service.Cost;
                    rowindex1++;
                }
                else if (Convert.ToDouble(service.Cost) > 250 && Convert.ToInt32(service.Cost) < 800)
                {
                    worksheet2.Cells[rowindex2, 1] = service.IdServices;
                    worksheet2.Cells[rowindex2, 2] = service.NameServices;
                    worksheet2.Cells[rowindex2, 3] = service.TypeOfService;
                    worksheet2.Cells[rowindex2, 4] = service.Cost;
                    rowindex2++;
                }
                else if (Convert.ToDouble(service.Cost) > 800)
                {
                    worksheet3.Cells[rowindex3, 1] = service.IdServices;
                    worksheet3.Cells[rowindex3, 2] = service.NameServices;
                    worksheet3.Cells[rowindex3, 3] = service.TypeOfService;
                    worksheet3.Cells[rowindex3, 4] = service.Cost;
                    rowindex3++;
                }
                else
                {

                }

            }
        }

        class Service
        {
            public int IdServices { get; set; }
            public string NameServices { get; set; }
            public string TypeOfService { get; set; }
            public string CodeService { get; set; }
            public int Cost { get; set; }

        }
        class Service1
        {
            public int IdServices { get; set; }
            public string NameServices { get; set; }
            public string TypeOfService { get; set; }
            public string CodeService { get; set; }
            public int Cost { get; set; }

        }
    }
}


