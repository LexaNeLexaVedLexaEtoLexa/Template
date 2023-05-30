using Makhov_4335;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4335.Windows.Makhov_4335
{
    /// <summary>
    /// Логика взаимодействия для ExcelPage.xaml
    /// </summary>
    public partial class ExcelPage : Page
    {
        public ExcelPage()
        {
            InitializeComponent();
            DBGridModel.ItemsSource = SecondaryFunctions.Services;
        }

        /// <summary>
        /// Реализация импорта данных в Базу Данных.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportBtn_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл"
            };
            if (!openFileDialog.ShowDialog() == true)
                return;

            ImportData(openFileDialog.FileName);

        }

        /// <summary>
        /// Реализация экспорта из Базы Данных в Excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportBtn_OnClick(object sender, RoutedEventArgs e) => ExportData();

        #region Методы импорта и экспорта

        /// <summary>
        /// Импорт данных.
        /// </summary>
        /// <param name="path"> Файл типа Excel. </param>
        private static void ImportData(string path)
        {
            try
            {
                var objWorkExcel = new Excel.Application();
                var objWorkBook = objWorkExcel.Workbooks.Open(path);
                var objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
                var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                var columns = lastCell.Column;
                var rows = lastCell.Row;
                var list = new string[rows, columns];
                for (var j = 0; j < columns; j++)
                {
                    for (var i = 1; i < rows; i++)
                    {
                        list[i, j] = objWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                objWorkBook.Close(false, Type.Missing, Type.Missing);
                objWorkExcel.Quit();
                GC.Collect();
                using (var db = new WordExcelDatabaseEntities())
                {
                    for (var i = 1; i < rows; i++)
                    {
                        var services = new Services
                        {
                            Id = list[i, 0],
                            ServiceName = list[i, 1],
                            ServiceType = list[i, 2],
                            ServiceCode = list[i, 3],
                            ServicePrice = list[i, 4]
                        };

                        db.Services.Add(services);
                    }

                    try
                    {
                        db.SaveChanges();
                        SecondaryFunctions.RefreshData();

                        MessageBox.Show("Данные импоритрованы успешно!",
                                        "Внимание!",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message,
                                        "Внимание",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Error);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                                "Внимание",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Экспорт данных.
        /// </summary>
        private static void ExportData()
        {
            #region Категории по возрасту

            var firstRangePrice = new List<Services>();
            var secondRangePrice = new List<Services>();
            var thirdRangePrice = new List<Services>();
            var categoriesPriceCount = 3;

            #endregion

            using (var db = new WordExcelDatabaseEntities())
            {
                #region Сортировка по цене

                firstRangePrice = db.Services.ToList()
                                             .Where(fR => Convert.ToInt32(fR.ServicePrice) >= 0 &&
                                                          Convert.ToInt32(fR.ServicePrice) <= 350)
                                             .GroupBy(fR => fR.ServiceName)
                                             .SelectMany(fR => fR)
                                             .ToList();

                secondRangePrice = db.Services.ToList()
                                              .Where(sR => Convert.ToInt32(sR.ServicePrice) >= 250 &&
                                                         Convert.ToInt32(sR.ServicePrice) <= 800)
                                              .GroupBy(sR => sR.ServiceName)
                                              .SelectMany(sR => sR)
                                              .ToList();

                thirdRangePrice = db.Services.ToList()
                                             .Where(tR => Convert.ToInt32(tR.ServicePrice) >= 800)
                                             .GroupBy(tR => tR.ServiceName)
                                             .SelectMany(tR => tR)
                                             .ToList();

                #endregion

                #region Создание Excel

                var app = new Excel.Application { SheetsInNewWorkbook = categoriesPriceCount };
                var book = app.Workbooks.Add(Type.Missing);

                #endregion

                #region Создание трех листов в Excel

                var startRowIndex = 1;
                var sheet1 = app.Worksheets.Item[1];
                sheet1.Name = "От 0 до 350 рублей";
                var sheet2 = app.Worksheets.Item[2];
                sheet2.Name = "От 250 до 800 рублей";
                var sheet3 = app.Worksheets.Item[3];
                sheet3.Name = "От 800 рублей";

                #endregion

                #region Создание колонок

                sheet1.Cells[1][startRowIndex] = "Id";
                sheet1.Cells[2][startRowIndex] = "Наименование услуги";
                sheet1.Cells[3][startRowIndex] = "Вид услуги";
                sheet1.Cells[4][startRowIndex] = "Стоимость";

                sheet2.Cells[1][startRowIndex] = "Id";
                sheet2.Cells[2][startRowIndex] = "Наименование услуги";
                sheet2.Cells[3][startRowIndex] = "Вид услуги";
                sheet2.Cells[4][startRowIndex] = "Стоимость";

                sheet3.Cells[1][startRowIndex] = "Id";
                sheet3.Cells[2][startRowIndex] = "Наименование услуги";
                sheet3.Cells[3][startRowIndex] = "Вид услуги";
                sheet3.Cells[4][startRowIndex] = "Стоимость";
                startRowIndex++;

                #endregion

                #region Заполнение таблицы данными

                #region Заполнение первого листа

                for (var i = 0; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in firstRangePrice)
                    {
                        sheet1.Cells[1][startRowIndex] = item.Id;
                        sheet1.Cells[2][startRowIndex] = item.ServiceName;
                        sheet1.Cells[3][startRowIndex] = item.ServiceType;
                        sheet1.Cells[4][startRowIndex] = item.ServicePrice;
                        startRowIndex++;
                    }

                }

                #endregion

                #region Заполнение второго листа

                for (var i = 1; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in secondRangePrice)
                    {
                        sheet2.Cells[1][startRowIndex] = item.Id;
                        sheet2.Cells[2][startRowIndex] = item.ServiceName;
                        sheet2.Cells[3][startRowIndex] = item.ServiceType;
                        sheet2.Cells[4][startRowIndex] = item.ServicePrice;
                        startRowIndex++;
                    }
                }

                #endregion

                #region Заполнение третьего листа

                for (var i = 2; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in thirdRangePrice)
                    {
                        sheet3.Cells[1][startRowIndex] = item.Id;
                        sheet3.Cells[2][startRowIndex] = item.ServiceName;
                        sheet3.Cells[3][startRowIndex] = item.ServiceType;
                        sheet3.Cells[4][startRowIndex] = item.ServicePrice;
                        startRowIndex++;
                    }
                }

                #endregion

                #endregion

                app.Visible = true;
            }
        }

        #endregion
    }
}
